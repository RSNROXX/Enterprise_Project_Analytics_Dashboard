import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from collections import defaultdict

from django.shortcuts import render, redirect, get_object_or_404            # type: ignore
from django.contrib import messages                                         # type: ignore
from django.http import HttpResponse                                        # type: ignore
from django.db.models import Q                                              # type: ignore

from .forms import UploadFileForm
from .models import Project, Metric, Department, UserGroup
from .constants import EXCEL_COL_MAP, ROLE_CONFIG, DEPT_PEOPLE_MAP, REPORT_ORDER_CONFIG, COMMON_REPORT_COLS

# ==============================================================================
# SECTION 1: GLOBAL HELPER SERVICES (Business Logic)
# ==============================================================================
def _handle_threshold_session(request):
    """
        Captures threshold changes from URL and saves them to Session.
        Returns a dictionary of {field_name: effective_value} merging Session > DB Defaults.
    """
    # 1. Initialize session storage if missing
    if 'threshold_overrides' not in request.session:
        request.session['threshold_overrides'] = {}

    # 2. Check for Reset Flag
    if request.GET.get('reset_thresholds'):
        request.session['threshold_overrides'] = {}
        request.session.modified = True
    
    # 3. Capture new changes from GET params
    # Format expected: 'thresh_pre_field_name' or 'thresh_post_field_name'
    for key, value in request.GET.items():
        if key.startswith('thresh_') and value:
            try:
                # Extract clean field name (remove prefix 'thresh_pre_' or 'thresh_post_')
                # We assume field names don't overlap between stages for simplicity, 
                # or we just map by field name since fields are unique in models mostly.
                parts = key.split('_', 2) # thresh, pre/post, field_name
                if len(parts) >= 3:
                    field_name = parts[2]
                    request.session['threshold_overrides'][field_name] = float(value)
                    request.session.modified = True
            except ValueError:
                continue
    # 4. Build Final Map (Session > DB Default)
    # We fetch all metrics to get defaults
    final_map = {}
    all_metrics = Metric.objects.all()
    for m in all_metrics:
        # Default to the DB 'min_threshold' if no session override exists
        # If your model hasn't been migrated yet, use 0.0 as fallback
        db_min = getattr(m, 'min_threshold', 0.0) 
        final_map[m.field_name] = request.session['threshold_overrides'].get(m.field_name, db_min)
    
    return final_map

def _get_request_params(request):
    """ 
        Standardizes extraction of Date Ranges, SBUs, and View Modes using SESSION PERSISTENCE.
    """
    # 1. Define Defaults
    default_start = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d')
    default_end = datetime.now().strftime('%Y-%m-%d')
    
    # USER REQUEST: Default SBU set to SME's only
    # If no filter is selected, we default to this specific list
    sme_defaults = ['Central', 'North', 'South', 'West']

    # 2. DATE LOGIC (Priority: URL > Session > Default)
    if request.GET.get('start'):
        start_str = request.GET.get('start')
        request.session['filter_start'] = start_str 
    else:
        start_str = request.session.get('filter_start', default_start) 


    if request.GET.get('end'):
        end_str = request.GET.get('end')
        request.session['filter_end'] = end_str 
    else:
        end_str = request.session.get('filter_end', default_end) 


    # 3. SBU LOGIC (Priority: URL > Session > Default)
    if 'sbu' in request.GET:
        sbu_filter = request.GET.getlist('sbu')
        request.session['filter_sbu'] = sbu_filter # Save new selection
    elif 'view' in request.GET: 
        # Switching departments? Keep previous selection or fall back to SME defaults
        sbu_filter = request.session.get('filter_sbu', sme_defaults)
    else:
        # Navigation/Fresh Load? Recall session or fall back to SME defaults
        sbu_filter = request.session.get('filter_sbu', sme_defaults)

    # 4. View Mode & Role 
    view_mode = request.GET.get('view', 'Sales') 
    role_filter = request.GET.get('metric_role') or request.GET.get('role') or 'All Roles'

    # 5. Process Dates
    try:
        start_dt = datetime.strptime(start_str, '%Y-%m-%d').date()
        end_dt = datetime.strptime(end_str, '%Y-%m-%d').date()
    except ValueError:
        start_dt = datetime.now().date() - timedelta(days=30)
        end_dt = datetime.now().date()

    roll_start = start_dt - timedelta(days=180)
    roll_end = end_dt + timedelta(days=240)

    return view_mode, start_str, end_str, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end

def _get_role_details(role_name):
    """ 
        Retrieves config from constants.py. Handles 'Sales - Sales Lead' format. 
    """
    simple_name = role_name.split(' - ')[-1] if ' - ' in role_name else role_name
    return ROLE_CONFIG.get(simple_name), simple_name

def _fetch_projects_filtered(sbu_filter, start_dt, end_dt, project_field=None):
    """ 
        Centralized Project Fetcher. 
        Applies SBU, Date Range, Role Filter, and CRITICAL Test Project Exclusion.
    """
    projects = Project.objects.filter(sbu__in=sbu_filter)

    # 1. Role Filter: If looking for specific role, exclude where missing
    if project_field:
        projects = projects.exclude(**{f"{project_field}__isnull": True})\
                           .exclude(**{f"{project_field}__exact": ""})

    # 2. Hard Exclusion for Bad Data / Test Projects
    projects = projects.exclude(project_code__isnull=True)\
                       .exclude(project_code__exact="")\
                       .exclude(project_code="PS-02AUG23-BB1_TEST-SOMERSET-01")

    # 3. Date Logic: Login Date OR Start Date must be in range
    projects = projects.filter(
        Q(login_date__range=[start_dt, end_dt]) | 
        Q(start_date__range=[start_dt, end_dt])
    ).distinct()
    
    return projects

def _get_scoring_engine_context(user_group, threshold_map):
    """ 
        Prepares Metrics & Weights for Scoring. 
    """
    metrics = Metric.objects.filter(
        metricweight__user_group=user_group, 
        metricweight__factor__gt=0
    ).distinct()
    
    stage_totals = {}
    valid_metrics = []

    for m in metrics:
        # We calculate the "Total Possible" based on the Max Threshold (Cap)
        # Assuming Max Threshold IS the max credits possible for that metric
        db_min = getattr(m, 'min_threshold', 1.0)
        db_max = getattr(m, 'max_threshold', 10.0)

        max_points = db_max
        
        stage_totals[m.stage] = stage_totals.get(m.stage, 0) + max_points

        # DASHBOARD OVERRIDE LOGIC:
        # If user changed threshold in Dashboard, it overrides the MINIMUM threshold.
        # Default falls back to DB min_threshold.
        effective_min = threshold_map.get(m.field_name, db_min)

        valid_metrics.append({
            'field': m.field_name, 
            'label': m.label,
            'stage': m.stage, 
            'min': effective_min,          
            'max': db_max,        
            'weight_factor': max_points    
        })
    return valid_metrics, stage_totals

def _calculate_project_score(project, valid_metrics, stage_totals):
    """ 
        Logic: 
      - If Actual < Min: 0 Pts
      - If Min <= Actual <= Max: (Actual - Min) Pts
      - If Actual > Max: Max Pts (Capped)
    """
    raw_stage = str(project.stage).strip().lower()
    current_stage = 'Post' if any(x in raw_stage for x in ['post', 'exec', 'ops', 'handover']) else 'Pre'

    total_possible = stage_totals.get(current_stage, 0)
    earned_points = 0.0
    
    if total_possible > 0:
        for vm in valid_metrics:
            if vm['stage'] == current_stage:
                val = getattr(project, vm['field'], 0.0)
                
                # 1. Calculate the Range Span
                # (Handle edge case where Max might be 0 to avoid division by zero)
                if vm['max'] > 0:
                    range_span = (vm['max'] - vm['min']) + 1
                    factor = range_span / vm['max']
                    
                    # 2. Calculate Points
                    calculated_points = val * factor
                    
                    # 3. Cap at Max Threshold? 
                    # Cap points at the Max Threshold to prevent "infinite" scores.
                    # We will CAP it at vm['max'] to be safe.
                    final_points = min(calculated_points, vm['max'])
                else:
                    final_points = 0.0
                
                earned_points += final_points
                
    return round(earned_points, 1), current_stage

def group_roles_by_dept(flat_roles):
    """
        Helper: Groups a list of role names into specific Departments in a specific order for Dropdown menus.
    """
    groups = {
        'Sales': ['Sales Head', 'Sales Lead'],
        'Design': ['DH', 'DM', 'ID', '3D'],
        'Operations': ['Cluster/BU Head', 'SPM/PM', 'SOM/OM', 'SS', 'CSC', 'MEP'],
        'Purchase': ['Purchase Head', 'Purchase Manager', 'Purchase Executive'],
        'Marketing': ['Marketing Head', 'Marketing Lead'],
        'Finance': ['Finance Head'],
    }
    
    # Create a result dict preserving the order above
    ordered_result = {k: [] for k in groups}
    ordered_result['Other'] = [] # Catch-all bucket

    for role in flat_roles:
        found = False
        for dept, role_list in groups.items():
            # Check if role matches exactly or contains the string
            if role in role_list:
                ordered_result[dept].append(role)
                found = True
                break
        if not found:
            ordered_result['Other'].append(role)
            
    # Remove empty departments and return
    return {k: v for k, v in ordered_result.items() if v}

# ==============================================================================
# SECTION 2: DASHBOARD SPECIFIC HELPERS
# ==============================================================================

def _fetch_metrics_from_db(view_mode, stage, role_filter, threshold_map):
    """ 
        Fetches dashboard metrics and determines primary/secondary visibility. 
    """
    try:
        dept = Department.objects.get(name__iexact=view_mode)
    except Department.DoesNotExist:
        return []

    metrics_qs = Metric.objects.filter(department=dept, stage=stage)\
                               .prefetch_related('visible_to_groups', 'metricweight_set__user_group')

    metrics_list = []
    for m in metrics_qs:
        # Combine Legacy Groups + Weighted Groups
        m2m_groups = {g.name for g in m.visible_to_groups.all()}
        weight_groups = {w.user_group.name for w in m.metricweight_set.all() if w.factor > 0}
        allowed_groups = list(m2m_groups | weight_groups)

        db_min = getattr(m, 'min_threshold', 0.0)
        effective_val = threshold_map.get(m.field_name, db_min)
        
        metrics_list.append({
            'label': m.label,
            'field': m.field_name,
            'def': effective_val,
            'success_cat': m.success_metric.name if m.success_metric else None,
            'success_color': m.success_metric.color if m.success_metric else 'secondary', 
            'allowed_groups': allowed_groups, 
            'id': m.pk 
        })
    return metrics_list

def _apply_people_filters(queryset, view_mode, request):
    """ 
        Dynamic Filtering based on View Mode (e.g. Sales Head filter). 
    """
    def _filter(qs, db_field, get_param):
        selected = request.GET.getlist(get_param)
        if not selected: return qs
        q = Q()
        for name in selected: q |= Q(**{f"{db_field}__icontains": name})
        return qs.filter(q)

    # Use DEPT_PEOPLE_MAP to determine which fields to filter for this view
    if view_mode in DEPT_PEOPLE_MAP:
        for db_field, label in DEPT_PEOPLE_MAP[view_mode]:
            # Construct the f_ prefix parameter key based on the db_field name
            # Logic: 'sales_head' -> 'f_s_head', 'ops_pm' -> 'f_o_pm'
            parts = db_field.split('_')
            if len(parts) == 2:
                prefix = parts[0][0] # 'sales' -> 's'
                suffix = parts[1]    # 'head' -> 'head'
                param = f"f_{prefix}_{suffix}"
                queryset = _filter(queryset, db_field, param)
    return queryset

def _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end):
    """ 
        Splits projects into Pre/Post buckets. 
    """
    q_rolling = Q(start_date__gte=roll_start) & Q(start_date__lte=end_dt) & \
                Q(end_date__gte=start_dt) & Q(end_date__lte=roll_end)

    qs_pre = Project.objects.none()
    qs_post = Project.objects.none()

    if view_mode == 'Sales':
        qs_pre = projects.filter(login_date__gte=start_dt, login_date__lte=end_dt, stage='Pre Sales')
        qs_post = projects.filter(q_rolling, stage='Post Sales')
    elif view_mode == 'Design':
        qs_pre = projects.filter(login_date__gte=start_dt, login_date__lte=end_dt)
        qs_post = projects.filter(q_rolling).distinct()
    elif view_mode == 'Operations':
        qs_post = projects.filter(q_rolling).distinct()
    
    return qs_pre, qs_post

def _get_dropdown_context(request):
    """ 
        Populates filter dropdowns with unique values from DB. 
    """
    def get_opts(field):
        try:
            vals = Project.objects.exclude(project_code="PS-02AUG23-BB1_TEST-SOMERSET-01").values_list(field, flat=True).distinct()
            return sorted([v for v in vals if v and str(v).strip()])
        except: return []

    people_opts = {
        'm_head': get_opts('m_head'), 'm_lead': get_opts('m_lead'),
        's_head': get_opts('sales_head'), 's_lead': get_opts('sales_lead'),
        'd_dh': get_opts('design_dh'), 'd_dm': get_opts('design_dm'),
        'd_id': get_opts('design_id'), 'd_3d': get_opts('design_3d'),
        'o_head': get_opts('ops_head'), 'o_pm': get_opts('ops_pm'),
        'o_om': get_opts('ops_om'), 'o_ss': get_opts('ops_ss'),
        'o_mep': get_opts('ops_mep'), 'o_csc': get_opts('ops_csc'),
        'p_head': get_opts('p_head'), 'p_exec': get_opts('p_exec'), 'p_mgr': get_opts('p_mgr'),
        'f_head': get_opts('f_head'),
    }
    
    selected_filters = {
        's_head': request.GET.getlist('f_s_head'), 's_lead': request.GET.getlist('f_s_lead'),
        'd_dh': request.GET.getlist('f_d_dh'), 'd_dm': request.GET.getlist('f_d_dm'),
        'd_id': request.GET.getlist('f_d_id'), 'd_3d': request.GET.getlist('f_d_3d'),
        'o_head': request.GET.getlist('f_o_head'), 'o_pm': request.GET.getlist('f_o_pm'),
        'o_om': request.GET.getlist('f_o_om'), 'o_ss': request.GET.getlist('f_o_ss'),
        'o_mep': request.GET.getlist('f_o_mep'), 'o_csc': request.GET.getlist('f_o_csc'),
    }
    return people_opts, selected_filters

# ==============================================================================
# SECTION 3: DASHBOARD VIEW
# ==============================================================================

def dashboard_view(request):
    threshold_map = _handle_threshold_session(request)

    view_mode, start_str, end_str, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    people_opts, selected_filters = _get_dropdown_context(request)
    
    # 1. Fetch Projects (Exclude Test Project)
    projects = Project.objects.filter(sbu__in=sbu_filter).exclude(project_code="PS-02AUG23-BB1_TEST-SOMERSET-01")
    projects = _apply_people_filters(projects, view_mode, request)
    all_departments = Department.objects.values_list('name', flat=True).order_by('name')

    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)
    pre_count = qs_pre.count()
    post_count = qs_post.count()

    def calculate_card_metrics(queryset, metrics_list, prefix):
        results_prim, results_sec = [], []
        
        for m in metrics_list:
            param_name = f"thresh_{prefix}_{m['field']}"
            threshold = m['def']
            user_input = request.GET.get(param_name)
            try: threshold = float(user_input) if user_input else m['def']
            except: threshold = m['def']

            filtered_qs = queryset.filter(**{f"{m['field']}__gte": threshold})
            count = filtered_qs.count()

            proj_data = list(filtered_qs.values('id', 'project_name', 'project_code').order_by('project_name'))

            item = {
                'label': m['label'], 'param': param_name, 'threshold': threshold, 
                'count': count, 'field': m['field'], 
                'success_cat': m['success_cat'], 
                'success_color': m['success_color'], 
                'project_list': proj_data
            }

            # Filter Logic: Is this card primary for the selected role?
            is_primary = False
            if role_filter == "All Roles":
                is_primary = True
            elif m['allowed_groups']:
                r_clean = str(role_filter).lower().strip()
                for group_name in m['allowed_groups']:
                    g_clean = str(group_name).lower().strip()
                    if r_clean in g_clean or g_clean in r_clean:
                        is_primary = True
                        break
            
            if is_primary: results_prim.append(item)
            else: results_sec.append(item)
                
        return results_prim, results_sec

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter, threshold_map)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter, threshold_map)

    pre_prim, pre_sec = [], []
    if view_mode != 'Operations':
        pre_prim, pre_sec = calculate_card_metrics(qs_pre, pre_metrics_db, 'pre')

    post_prim, post_sec = [], []
    post_prim, post_sec = calculate_card_metrics(qs_post, post_metrics_db, 'post')

    sbu_opts = sorted([s for s in Project.objects.values_list('sbu', flat=True).distinct() if s])

    has_overrides = bool(request.session.get('threshold_overrides'))
    
    context = {
        'view_mode': view_mode, 'start_date': start_str, 'end_date': end_str,
        'sbus': sbu_opts or ['North', 'South', 'West', 'Central'], 'selected_sbus': sbu_filter,
        'current_role': role_filter, 'people_opts': people_opts, 'selected_filters': selected_filters,
        'pre_prim': pre_prim, 'pre_sec': pre_sec, 'pre_count': pre_count,
        'post_prim': post_prim, 'post_sec': post_sec, 'post_count': post_count,
        'all_departments': all_departments,
        'has_overrides': has_overrides,
    }
    return render(request, 'core/dashboard.html', context)

# ==============================================================================
# SECTION 3: UPLOAD LOGIC (FULLY RESTORED)
# ==============================================================================

def upload_view(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                file = request.FILES['file']
                xls = pd.ExcelFile(file)
                sheet_map = {'sales': '', 'design': '', 'operation': ''}
                
                # Find sheets (Case insensitive)
                for name in xls.sheet_names:
                    lower = str(name).lower()
                    if 'sales' in lower: sheet_map['sales'] = str(name)
                    elif 'design' in lower: sheet_map['design'] = str(name)
                    elif 'operation' in lower or 'ops' in lower: sheet_map['operation'] = str(name)

                project_data_map = {}

                # Helpers
                def clean_str(val): 
                    if pd.isnull(val): return ''
                    s = str(val).strip()
                    return '' if s.lower() == 'nan' else s

                def parse_date(val): 
                    return pd.to_datetime(val).date() if pd.notnull(val) else None

                def clean_num(val):
                    if pd.isnull(val): return 0.0
                    try: return float(str(val).replace('%','').replace(',','').strip())
                    except: return 0.0
                
                def get_clean_id(row_dict):
                    for c in ['project code','code','lead id']:
                        if c in row_dict and row_dict[c]:
                            v = str(row_dict[c]).strip().upper()
                            if v and v != 'NAN': return v.replace('.0','')
                    return None

                def process_sheet(sheet_name, sheet_type):
                    if not sheet_name: return
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    
                    # 1. FORCE LOWERCASE HEADERS
                    df.columns = [str(c).strip().lower() for c in df.columns]

                    # 2. CALCULATED COLUMNS
                    if sheet_type == 'design':
                        if 'no key plans spaces' in df.columns and 'mapped spaces' in df.columns:
                            df['key plans ratio'] = pd.to_numeric(df['no key plans spaces'], errors='coerce').fillna(0) / pd.to_numeric(df['mapped spaces'], errors='coerce').replace(0,1).fillna(0)
                        if 'layouts' in df.columns and 'furniture layouts' in df.columns:
                            df['other layouts'] = df['layouts'] - df['furniture layouts']
                    
                    if sheet_type == 'operation':
                        ops_calcs = [
                            ('wpr half week','wpr download weeks','weeks till date'), 
                            ('manpower ratio','actual manpower','planned manpower'),
                            ('dpr ratio','dpr added days','days till date'), 
                            ('manpower day ratio','manpower added days','days till date')
                        ]
                        for t, n, d in ops_calcs:
                            if n in df.columns and d in df.columns:
                                df[t] = (pd.to_numeric(df[n], errors='coerce') / pd.to_numeric(df[d], errors='coerce').replace(0,1)).fillna(0)

                    # 3. ROW ITERATION
                    for _, row_series in df.iterrows():
                        row = {k: v for k, v in row_series.items()}
                        p_id = get_clean_id(row)
                        if not p_id: continue

                        if p_id not in project_data_map: 
                            project_data_map[p_id] = {'project_code': p_id}
                        
                        # Metadata Mapping
                        meta_config = {
                            'project_name': ['project name', 'name'],
                            'sbu':          ['sbu', 'region'],
                            'stage':        ['stage', 'status'],
                            'floors':       ['floors', 'no of floors'],
                            'project_type': ['project type', 'type'],
                            'lead_id':      ['lead id', 'lead', 'id'],

                            'sales_head':   ['sales head', 's head'],
                            'sales_lead':   ['sales lead', 's lead'],

                            'design_dh':    ['dh', 'design head'],
                            'design_dm':    ['dm', 'design lead', 'design manager'],
                            'design_id':    ['id', 'design id'],
                            'design_3d':    ['3d', '3d visualizer'],

                            'ops_head':     ['cluster/bu head', 'ops head'],
                            'ops_pm':       ['spm/pm', 'project manager', 'pm'],
                            'ops_om':       ['som/om', 'ops manager', 'om'],
                            'ops_ss':       ['ss', 'site supervisor'],
                            'ops_mep':      ['mep'],
                            'ops_csc':      ['csc']
                        }

                        for db_field, options in meta_config.items():
                            for opt in options:
                                if opt in row:
                                    val = clean_str(row[opt])
                                    if val: 
                                        project_data_map[p_id][db_field] = val
                                    break 

                        # Dates
                        date_map = {'login_date': 'project login date', 'start_date': 'project start date', 'end_date': 'project end date'}
                        for db, xl in date_map.items():
                            if xl in row:
                                v = parse_date(row[xl])
                                if v: project_data_map[p_id][db] = v

                        # Metrics
                        full_map = EXCEL_COL_MAP.copy()
                        full_map.update({
                            'Key Plans Ratio':'key_plans_ratio', 'Other Layouts':'other_layouts', 
                            'WPR Half Week':'wpr_half_week', 'Manpower Ratio':'manpower_ratio', 
                            'DPR Ratio':'dpr_ratio', 'Manpower Day Ratio':'manpower_day_ratio'
                        })

                        for xl_col, db_field in full_map.items():
                            xl_lower = str(xl_col).strip().lower()
                            if xl_lower in row:
                                v = clean_num(row[xl_lower])
                                if v != 0 or db_field not in project_data_map[p_id]:
                                    project_data_map[p_id][db_field] = v

                process_sheet(sheet_map['sales'], 'sales')
                process_sheet(sheet_map['design'], 'design')
                process_sheet(sheet_map['operation'], 'operation')

                if project_data_map:
                    Project.objects.all().delete()
                    Project.objects.bulk_create([Project(**d) for d in project_data_map.values()])
                    messages.success(request, f"Restored {len(project_data_map)} projects. Database updated.")
                else:
                    messages.error(request, "No valid project data found in file.")
                
                return redirect('dashboard')

            except Exception as e:
                messages.error(request, f"Upload Failed: {str(e)}")
    else:
        form = UploadFileForm()
    return render(request, 'core/upload.html', {'form': form})

def project_detail(request, pk):
    project = get_object_or_404(Project, pk=pk)
    return render(request, 'core/project_detail.html', {'project': project})

# ==============================================================================
# SECTION 4: EXPORTS & REPORTS (FIXED: SEPARATED SUMMARY & DETAILED)
# ==============================================================================

def export_view(request):
    """ 
        Excel Export for Leadership Summary (Counts & %) 
    """
    return _handle_summary_report(request, is_excel=True)

def report_view(request):
    """ 
        HTML Live View for Leadership Summary (Counts & %) 
    """
    return _handle_summary_report(request, is_excel=False)

def export_detailed_view(request):
    """ 
        Excel Export for Detailed Project List (Rows & Columns) 
    """
    return _handle_detailed_report(request, is_excel=True)

def report_detailed_view(request):
    """ 
        HTML Live View for Detailed Project List (Rows & Columns) 
    """
    return _handle_detailed_report(request, is_excel=False)

# --- INTERNAL REPORT HANDLERS ---

def _handle_summary_report(request, is_excel=False):
    threshold_map = _handle_threshold_session(request) 
    view_mode, _, _, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    projects = Project.objects.filter(sbu__in=sbu_filter).exclude(project_code="PS-02AUG23-BB1_TEST-SOMERSET-01")
    projects = _apply_people_filters(projects, view_mode, request)
    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter, threshold_map)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter, threshold_map)

    def generate_summary_df(queryset, metrics_list, prefix):
        total = queryset.count()
        data = [{"Metric Name": "TOTAL PROJECTS", "Threshold": "-", "Value": total, "%": "-"}]
        for m in metrics_list:
            threshold = m['def']
            param_name = f"thresh_{prefix}_{m['field']}"
            user_input = request.GET.get(param_name)
            count = queryset.filter(**{f"{m['field']}__gte": threshold}).count()
            pct = round((count / total * 100), 1) if total > 0 else 0.0
            data.append({"Metric Name": m['label'], "Category": m['success_cat'], "Threshold": threshold, "Value": count, "%": f"{pct}%"})
        return pd.DataFrame(data)

    df_pre = generate_summary_df(qs_pre, pre_metrics_db, 'pre') if view_mode != 'Operations' else pd.DataFrame()
    df_post = generate_summary_df(qs_post, post_metrics_db, 'post')

    if is_excel:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if not df_pre.empty: df_pre.to_excel(writer, sheet_name='Pre-Stage', index=False)
            if not df_post.empty: df_post.to_excel(writer, sheet_name='Post-Stage', index=False)
            if df_pre.empty and df_post.empty: pd.DataFrame({'Info': ['No Data']}).to_excel(writer, sheet_name='Empty', index=False)
        buffer.seek(0)
        response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="Summary_{view_mode}.xlsx"'
        return response
    else:
        context = {
            'view_mode': view_mode, 'start_date': str(start_dt), 'end_date': str(end_dt),
            'selected_sbus': sbu_filter, 
            'df_pre': df_pre.to_html(classes='table table-striped table-hover', index=False, justify='left') if not df_pre.empty else None,
            'df_post': df_post.to_html(classes='table table-striped table-hover', index=False, justify='left') if not df_post.empty else None,
        }
        return render(request, 'core/report.html', context)

def _handle_detailed_report(request, is_excel=False):
    threshold_map = _handle_threshold_session(request)
    view_mode, _, _, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    projects = Project.objects.filter(sbu__in=sbu_filter).exclude(project_code="PS-02AUG23-BB1_TEST-SOMERSET-01")
    projects = _apply_people_filters(projects, view_mode, request)
    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)

    def generate_df(queryset, metrics_list, stage_key):
        if not metrics_list and not queryset.exists(): return pd.DataFrame()

        role_cols = []
        if view_mode in DEPT_PEOPLE_MAP:
            role_cols = [field for field, label in DEPT_PEOPLE_MAP[view_mode]]

        std_cols = ['project_name', 'project_code', 'lead_id', 'floors', 'project_type','sbu', 'stage', ]
        metric_fields = [m['field'] for m in metrics_list]
        
        fetch_fields = std_cols + role_cols + metric_fields
        data = list(queryset.values(*fetch_fields))
        
        if not data: return pd.DataFrame()
        df = pd.DataFrame(data)
        
        rename_map = {m['field']: m['label'] for m in metrics_list}
        if view_mode in DEPT_PEOPLE_MAP:
            for field, label in DEPT_PEOPLE_MAP[view_mode]:
                rename_map[field] = label
        
        # Can be used to rename metadata fields back into readable format. 
        rename_map.update({
            'project_name': 'Project Name',
            'project_code': 'Project Code',
            'lead_id': 'Lead ID',
            'floors': 'Floors',
            'project_type': 'Project Type',
            'stage': 'Stage',
            'sbu': 'SBU',
        })

        df = df.rename(columns=rename_map).fillna('')

        dept_config = REPORT_ORDER_CONFIG.get(view_mode, COMMON_REPORT_COLS)

        if isinstance(dept_config, dict):
            preferred_order = dept_config.get(stage_key, COMMON_REPORT_COLS)
        else:
            preferred_order = dept_config

        final_cols = [c for c in preferred_order if c in df.columns]
        extra_cols = [c for c in df.columns if c not in final_cols]

        return df[final_cols + extra_cols]

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter, threshold_map)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter, threshold_map)
    df_pre = generate_df(qs_pre, pre_metrics_db, 'Pre') if view_mode != 'Operations' else pd.DataFrame()
    df_post = generate_df(qs_post, post_metrics_db, 'Post')

    if is_excel:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if not df_pre.empty: df_pre.to_excel(writer, sheet_name='Pre-Stage', index=False)
            if not df_post.empty: df_post.to_excel(writer, sheet_name='Post-Stage', index=False)
        buffer.seek(0)
        response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="Detailed_{view_mode}.xlsx"'
        return response
    else:
        context = {
            'view_mode': view_mode, 'start_date': str(start_dt), 'end_date': str(end_dt),
            
            # PRE-STAGE DATA
            'pre_columns': df_pre.columns.tolist() if not df_pre.empty else [],
            'pre_data': df_pre.to_dict('records') if not df_pre.empty else [],
            
            # POST-STAGE DATA
            'post_columns': df_post.columns.tolist() if not df_post.empty else [],
            'post_data': df_post.to_dict('records') if not df_post.empty else [],
        }
        return render(request, 'core/report_detailed.html', context)

# ==============================================================================
# SECTION 5: NEW FEATURES (Leaderboard & Scorecard)
# ==============================================================================

def project_scorecard_view(request, project_code):
    threshold_map = _handle_threshold_session(request)
    project = get_object_or_404(Project, project_code=project_code)
    
    raw_role_param = request.GET.get('metric_role', 'Design - ID')
    search_term = raw_role_param.split(' - ')[-1] if ' - ' in raw_role_param else raw_role_param
    
    user_group = UserGroup.objects.filter(name__icontains=search_term).first()
    all_groups = UserGroup.objects.select_related('department').order_by('department__name', 'name')
    
    if not user_group:
         return render(request, 'core/project_scorecard.html', {
            'project': project, 'error': f"Role '{search_term}' not found.",
            'all_groups': all_groups, 'selected_role_full': raw_role_param 
        })

    valid_metrics, stage_totals = _get_scoring_engine_context(user_group, threshold_map)
    project_score, metric_stage_key = _calculate_project_score(project, valid_metrics, stage_totals)
    
    total_factor_sum = stage_totals.get(metric_stage_key, 0)
    final_scores = []
    
    COLOR_SUCCESS = "#10b981" 
    COLOR_WARNING = "#f59e0b"
    COLOR_DANGER  = "#ef4444"

    for vm in valid_metrics:
        if vm['stage'] == metric_stage_key:
            current_value = getattr(project, vm['field'], 0.0)
            
            # 1. Scoring Logic 
            if vm['max'] > 0:
                range_span = (vm['max'] - vm['min']) + 1
                factor = range_span / vm['max']
                calculated_points = current_value * factor
                points = min(calculated_points, vm['max'])
            else:
                points = 0.0
            
            # 2. VISUAL BAR LOGIC 
            # Rule: Bar always covers [Min, Max]. 
            # If Value is OUTSIDE this range, we stretch the bar to include it.

            # Default boundaries
            display_start = vm['min']
            display_end = vm['max']

            # Adjust if value is lower than Min
            if current_value < vm['min']:
                display_start = current_value
            
            # Adjust if value is higher than Max
            if current_value > vm['max']:
                display_end = current_value

            # Calculate Marker % relative to this Dynamic Display Range
            total_span = display_end - display_start
            if total_span == 0:
                marker_pct = 100 if current_value > 0 else 0
            else:
                marker_pct = ((current_value - display_start) / total_span) * 100
            
            # Progress % for the "Status Text" (Standard 0-100% of Max Points)
            progress_pct = (points / vm['max'] * 100) if vm['max'] > 0 else 0
            
            # --- 3. COLORS ---
            if points >= vm['max']: # Use >= in case cap is hit
                color = COLOR_SUCCESS
                icon = "fas fa-check-circle"
            elif points > 0:
                color = COLOR_WARNING
                icon = "fas fa-exclamation-circle"
            else:
                color = COLOR_DANGER
                icon = "fas fa-times-circle"

            final_scores.append({
                'metric': vm['label'], 
                'min': vm['min'], 
                'max': vm['max'],
                
                # Visual Data for the Bar
                'bar_start': int(display_start) if display_start % 1 == 0 else display_start,
                'bar_end': int(display_end) if display_end % 1 == 0 else display_end,
                'marker_pct': marker_pct,

                'actual': int(current_value) if current_value % 1 == 0 else round(current_value, 1),
                'points_earned': round(points, 1),
                'status_text': f"{int(progress_pct)}%",
                'color': color, 'icon': icon,
                'factor': vm['max'] 
            })

    final_scores.sort(key=lambda x: x['factor'], reverse=True)

    all_role_keys = sorted(ROLE_CONFIG.keys())
    grouped_roles = group_roles_by_dept(all_role_keys)

    # 2. CALCULATE PROJECT LEVEL COLOR
    project_color = COLOR_DANGER
    if project_score >= 80:
        project_color = COLOR_SUCCESS
    elif project_score >= 50:
        project_color = COLOR_WARNING

    # 3. CALCULATE METRIC LEVEL COLORS
    for score in final_scores:
        # Logic: Full points = Green, Partial = Orange, Zero = Red
        if score['points_earned'] == score['factor']:
            score['color'] = COLOR_SUCCESS
            score['bs_class'] = 'success' # For Bootstrap classes like bg-success
            score['icon'] = 'fas fa-check-circle'
        elif score['points_earned'] > 0:
            score['color'] = COLOR_WARNING
            score['bs_class'] = 'warning'
            score['icon'] = 'fas fa-exclamation-circle'
        else:
            score['color'] = COLOR_DANGER
            score['bs_class'] = 'danger'
            score['icon'] = 'fas fa-times-circle'

    return render(request, 'core/project_scorecard.html', {
        'project': project, 'user_group': user_group,
        'project_color': project_color, 'scores': final_scores,
        'selected_group_id': user_group.id, 'all_groups': all_groups, 
        'total_factor': total_factor_sum, 'scores': final_scores, 
        'project_total': project_score,
        'project_total_int': int(project_score),
        'metric_stage': metric_stage_key,
        'grouped_roles': grouped_roles,
        'metric_role': raw_role_param,
    })

def leaderboard_view(request):
    threshold_map = _handle_threshold_session(request)
    _, start_str, end_str, start_dt, end_dt, sbu_filter, _, _, _ = _get_request_params(request)
    
    all_sbu_options = list(Project.objects.exclude(sbu__isnull=True).exclude(sbu="").values_list('sbu', flat=True).distinct())
    all_sbu_options.sort()

    all_role_keys = sorted(ROLE_CONFIG.keys())

    grouped_roles = group_roles_by_dept(all_role_keys)

    selected_role_name = request.GET.get('role', 'Sales Lead')
    config, simple_role_name = _get_role_details(selected_role_name)
    
    if not config: return render(request, 'core/leaderboard.html', {'error': "Role not found."})

    project_field = config['field']
    user_group = UserGroup.objects.filter(name__icontains=simple_role_name).first()
    if not user_group: return render(request, 'core/leaderboard.html', {'error': "User Group config missing."})

    projects = _fetch_projects_filtered(sbu_filter, start_dt, end_dt, project_field)
    valid_metrics, stage_totals = _get_scoring_engine_context(user_group, threshold_map)

    leaderboard = {}
    for proj in projects:
        if not proj.project_code or not str(proj.project_code).strip(): continue
        user_email = getattr(proj, project_field)
        if not user_email: continue
        user_key = str(user_email).strip().lower()
        
        if user_key not in leaderboard:
            leaderboard[user_key] = {'name': user_email, 'total_score': 0, 'projects': 0, 'breakdown': []}

        project_score, stage_name = _calculate_project_score(proj, valid_metrics, stage_totals)

        leaderboard[user_key]['total_score'] += project_score
        leaderboard[user_key]['projects'] += 1
        leaderboard[user_key]['breakdown'].append({
            'project_name': proj.project_name or proj.project_code,
            'code': proj.project_code, 'stage': stage_name,
            'sbu': proj.sbu, 'score': project_score
        })

    sorted_leaderboard = sorted(leaderboard.values(), key=lambda x: x['total_score'], reverse=True)
    all_scores = [u['total_score'] for u in leaderboard.values()]
    total_users = len(all_scores)

    for idx, row in enumerate(sorted_leaderboard, 1):
        row['rank'] = idx
        row['breakdown'].sort(key=lambda x: x['score'], reverse=True)
        if total_users > 0:
            people_beaten = sum(1 for s in all_scores if s <= row['total_score'])
            row['percentile'] = int((people_beaten / total_users) * 100)
        else: row['percentile'] = 0
        row['total_score'] = round(row['total_score'], 1)

    context = {
        'leaderboard': sorted_leaderboard, 'selected_role': selected_role_name, 'grouped_roles': grouped_roles,
        'all_roles': sorted(ROLE_CONFIG.keys()), 'start_date': start_str, 'end_date': end_str,
        'selected_sbus': sbu_filter, 'sbu_options': all_sbu_options,
        'link_view': config['dept'], 'link_param': config['link']
    }
    return render(request, 'core/leaderboard.html', context)

def leaderboard_summary_view(request):
    threshold_map = _handle_threshold_session(request)
    _, start_str, end_str, start_dt, end_dt, sbu_filter, _, _, _ = _get_request_params(request)
    all_sbu_options = list(Project.objects.exclude(sbu__isnull=True).exclude(sbu="").values_list('sbu', flat=True).distinct())
    all_sbu_options.sort()
    sbu_filter = request.GET.getlist('sbu') or all_sbu_options or ['North', 'South', 'East', 'Central']

    hall_of_fame = defaultdict(dict)

    for role_name, config in ROLE_CONFIG.items():
        if 'dept' not in config: continue 
        department = config['dept']
        project_field = config['field']
        link_param = config['link']
        
        search_term = role_name.replace("Design ", "").replace("Ops ", "").replace("Sales ", "")
        user_group = UserGroup.objects.filter(name__icontains=search_term).first()
        if not user_group: continue

        projects = _fetch_projects_filtered(sbu_filter, start_dt, end_dt, project_field)
        valid_metrics, stage_totals = _get_scoring_engine_context(user_group, threshold_map)

        role_leaderboard = {}
        for proj in projects:
            user_email = getattr(proj, project_field)
            if not user_email: continue
            user_key = str(user_email).strip().lower()
            if user_key not in role_leaderboard:
                role_leaderboard[user_key] = {'name': user_email, 'total_score': 0}

            score, _ = _calculate_project_score(proj, valid_metrics, stage_totals)
            role_leaderboard[user_key]['total_score'] += score

        sorted_users = sorted(role_leaderboard.values(), key=lambda x: x['total_score'], reverse=True)
        top_two = sorted_users[:2]
        for user in top_two:
            user['total_score'] = int(round(user['total_score'], 0))
            user['link_param'] = link_param 

        if top_two: hall_of_fame[department][role_name] = top_two

    context = {
        'hall_of_fame': dict(hall_of_fame), 'start_date': start_str, 'end_date': end_str,
        'selected_sbus': sbu_filter, 'sbu_options': all_sbu_options,
    }
    return render(request, 'core/leaderboard_summary.html', context)