import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from django.shortcuts import render, redirect, get_object_or_404    #type:ignore
from django.contrib import messages                                 #type:ignore
from django.http import HttpResponse                                #type:ignore
from django.db.models import Q, Sum, Count, F                       #type:ignore

from .forms import UploadFileForm
from .models import Project, Metric, Department, UserGroup, MetricWeight
from .constants import EXCEL_COL_MAP    

# ==============================================================================
#  INTERNAL HELPER FUNCTIONS
# ==============================================================================

def _get_request_params(request):
    default_start = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d')
    default_end = datetime.now().strftime('%Y-%m-%d')
    
    view_mode = request.GET.get('view', 'Sales') 
    start_str = request.GET.get('start', default_start)
    end_str = request.GET.get('end', default_end)
    sbu_filter = request.GET.getlist('sbu') or ['North', 'South', 'West', 'Central']
    role_filter = request.GET.get('metric_role', 'All Roles')

    try:
        start_dt = datetime.strptime(start_str, '%Y-%m-%d').date()
        end_dt = datetime.strptime(end_str, '%Y-%m-%d').date()
    except ValueError:
        start_dt = datetime.now().date() - timedelta(days=30)
        end_dt = datetime.now().date()

    roll_start = start_dt - timedelta(days=180)
    roll_end = end_dt + timedelta(days=240)

    return view_mode, start_str, end_str, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end

def _fetch_metrics_from_db(view_mode, stage, role_filter):
    try:
        dept = Department.objects.get(name__iexact=view_mode)
    except Department.DoesNotExist:
        return []

    # FETCH EVERYTHING for this department & stage
    # UPDATE: Prefetch 'metricweight_set' so we can check weights too
    metrics_qs = Metric.objects.filter(department=dept, stage=stage)\
                               .prefetch_related('visible_to_groups', 'metricweight_set__user_group')

    metrics_list = []
    for m in metrics_qs:
        # 1. Get groups from the "Visibility Filter" box (Old way)
        m2m_groups = {g.name for g in m.visible_to_groups.all()}
        
        # 2. Get groups from the "Weight Table" (New way)
        # If you gave it a weight > 0, it IS visible to that group!
        weight_groups = {w.user_group.name for w in m.metricweight_set.all() if w.factor > 0}
        
        # 3. Combine them (Union) - This ensures NO duplicates
        allowed_groups = list(m2m_groups | weight_groups)
        
        metrics_list.append({
            'label': m.label,
            'field': m.field_name,
            'def': m.default_threshold,
            'success_cat': m.success_metric.name if m.success_metric else None,
            'success_color': m.success_metric.color if m.success_metric else 'secondary', 
            'allowed_groups': allowed_groups, # Now contains both sources
            'id': m.pk 
        })
    return metrics_list

def _apply_people_filters(queryset, view_mode, request):
    def _filter(qs, db_field, get_param):
        selected = request.GET.getlist(get_param)
        if not selected: return qs
        q = Q()
        for name in selected: q |= Q(**{f"{db_field}__icontains": name})
        return qs.filter(q)

    if view_mode == 'Sales':
        queryset = _filter(queryset, 'sales_head', 'f_s_head')
        queryset = _filter(queryset, 'sales_lead', 'f_s_lead')

    elif view_mode == 'Design':
        queryset = _filter(queryset, 'design_dh', 'f_d_dh')
        queryset = _filter(queryset, 'design_dm', 'f_d_dm')
        queryset = _filter(queryset, 'design_id', 'f_d_id')
        queryset = _filter(queryset, 'design_3d', 'f_d_3d')

    elif view_mode == 'Operations':
        queryset = _filter(queryset, 'ops_head', 'f_o_head')
        queryset = _filter(queryset, 'ops_pm', 'f_o_pm')
        queryset = _filter(queryset, 'ops_om', 'f_o_om')
        queryset = _filter(queryset, 'ops_ss', 'f_o_ss')
        queryset = _filter(queryset, 'ops_mep', 'f_o_mep')
        queryset = _filter(queryset, 'ops_csc', 'f_o_csc')
    return queryset

def _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end):
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
    def get_opts(field):
        try:
            vals = Project.objects.values_list(field, flat=True).distinct()
            return sorted([v for v in vals if v and str(v).strip()])
        except: return []

    people_opts = {
        'm_head': get_opts('m_head'),
        'm_lead': get_opts('m_lead'),

        's_head': get_opts('sales_head'), 
        's_lead': get_opts('sales_lead'),

        'd_dh': get_opts('design_dh'), 
        'd_dm': get_opts('design_dm'),
        'd_id': get_opts('design_id'), 
        'd_3d': get_opts('design_3d'),

        'o_head': get_opts('ops_head'), 
        'o_pm': get_opts('ops_pm'),
        'o_om': get_opts('ops_om'), 
        'o_ss': get_opts('ops_ss'),
        'o_mep': get_opts('ops_mep'), 
        'o_csc': get_opts('ops_csc'),

        'p_head': get_opts('p_head'),
        'p_exec': get_opts('p_exec'),
        'p_mgr': get_opts('p_mgr'),

        'f_head': get_opts('f_head'),
    }
    
    selected_filters = {
        's_head': request.GET.getlist('f_s_head'), 
        's_lead': request.GET.getlist('f_s_lead'),

        'd_dh': request.GET.getlist('f_d_dh'), 
        'd_dm': request.GET.getlist('f_d_dm'),
        'd_id': request.GET.getlist('f_d_id'), 
        'd_3d': request.GET.getlist('f_d_3d'),

        'o_head': request.GET.getlist('f_o_head'), 
        'o_pm': request.GET.getlist('f_o_pm'),
        'o_om': request.GET.getlist('f_o_om'), 
        'o_ss': request.GET.getlist('f_o_ss'),
        'o_mep': request.GET.getlist('f_o_mep'), 
        'o_csc': request.GET.getlist('f_o_csc'),
    }
    return people_opts, selected_filters

# ==============================================================================
# 1. UPLOAD VIEW
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

                # --- CLEANING FUNCTIONS ---
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
                    # Check common ID column names (lowercase)
                    for c in ['project code','code','lead id']:
                        if c in row_dict and row_dict[c]:
                            v = str(row_dict[c]).strip().upper()
                            if v and v != 'NAN': return v.replace('.0','')
                    return None

                def process_sheet(sheet_name, sheet_type):
                    if not sheet_name: return
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    
                    # 1. FORCE LOWERCASE HEADERS (Crucial Fix)
                    df.columns = [str(c).strip().lower() for c in df.columns]

                    # 2. CALCULATED COLUMNS (Updated for Lowercase Headers)
                    # We must check for lowercase versions of the columns
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
                    # We convert the row to a dictionary for easier access
                    for _, row_series in df.iterrows():
                        # Create a lowercase-key dictionary for this row
                        row = {k: v for k, v in row_series.items()}
                        
                        p_id = get_clean_id(row)
                        if not p_id: continue

                        if p_id not in project_data_map: 
                            project_data_map[p_id] = {'project_code': p_id}
                        
                        # --- A. METADATA & PEOPLE MAPPING (The Fuzzy Logic) ---
                        meta_config = {
                            # Meta
                            'project_name': ['project name', 'name'],
                            'sbu':          ['sbu', 'region'],
                            'stage':        ['stage', 'status'],
                            
                            # People
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
                                    break # Stop after finding the first match

                        # --- B. DATES ---
                        date_map = {'login_date': 'project login date', 'start_date': 'project start date', 'end_date': 'project end date'}
                        for db, xl in date_map.items():
                            if xl in row:
                                v = parse_date(row[xl])
                                if v: project_data_map[p_id][db] = v

                        # --- C. METRICS (The Fix for "0 Values") ---
                        # We iterate over EXCEL_COL_MAP, but we LOWERCASE the key first
                        full_map = EXCEL_COL_MAP.copy()
                        full_map.update({
                            'Key Plans Ratio':'key_plans_ratio', 'Other Layouts':'other_layouts', 
                            'WPR Half Week':'wpr_half_week', 'Manpower Ratio':'manpower_ratio', 
                            'DPR Ratio':'dpr_ratio', 'Manpower Day Ratio':'manpower_day_ratio'
                        })

                        for xl_col, db_field in full_map.items():
                            xl_lower = str(xl_col).strip().lower() # <--- THE FIX
                            
                            if xl_lower in row:
                                v = clean_num(row[xl_lower])
                                # Only overwrite if value > 0 or not yet set
                                if v != 0 or db_field not in project_data_map[p_id]:
                                    project_data_map[p_id][db_field] = v

                process_sheet(sheet_map['sales'], 'sales')
                process_sheet(sheet_map['design'], 'design')
                process_sheet(sheet_map['operation'], 'operation')

                # SAFE SAVE (Atomic Transaction recommended in production, but simplified here)
                if project_data_map:
                    Project.objects.all().delete()
                    Project.objects.bulk_create([Project(**d) for d in project_data_map.values()])
                    messages.success(request, f"Restored {len(project_data_map)} projects. Values should be back.")
                else:
                    messages.error(request, "No valid project data found in file.")
                
                return redirect('dashboard')

            except Exception as e:
                messages.error(request, f"Upload Failed: {str(e)}")
    else:
        form = UploadFileForm()
    return render(request, 'core/upload.html', {'form': form})

# ==============================================================================
# 2. DASHBOARD VIEW
# ==============================================================================
def dashboard_view(request):
    view_mode, start_str, end_str, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    people_opts, selected_filters = _get_dropdown_context(request)
    
    projects = Project.objects.filter(sbu__in=sbu_filter)
    projects = _apply_people_filters(projects, view_mode, request)
    all_departments = Department.objects.values_list('name', flat=True).order_by('name')

    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)
    pre_count = qs_pre.count()
    post_count = qs_post.count()

    def calculate_card_metrics(queryset, metrics_list, prefix):
        results_prim, results_sec = [], []
        
        for m in metrics_list:
            param_name = f"thresh_{prefix}_{m['field']}"
            user_input = request.GET.get(param_name)
            try: threshold = float(user_input) if user_input else m['def']
            except: threshold = m['def']

            filtered_qs = queryset.filter(**{f"{m['field']}__gte": threshold})
            count = filtered_qs.count()
            proj_data = list(filtered_qs.values('id', 'project_name').order_by('project_name'))

            item = {
                'label': m['label'], 'param': param_name, 'threshold': threshold, 
                'count': count, 'field': m['field'], 
                'success_cat': m['success_cat'], 
                'success_color': m.get('success_color', 'secondary'), 
                'project_list': proj_data
            }

            # --- ROBUST FILTER LOGIC ---
            is_primary = False
            
            if role_filter == "All Roles":
                is_primary = True
            elif not m['allowed_groups']:
                # If NO weights/visibility set, default to Secondary (Safety Net)
                is_primary = False 
            else:
                # Fuzzy Match: Check if "Sales Lead" is in "Sales - Sales Lead"
                # OR if "ID" is in "Design - ID"
                # This ensures partial matches work
                r_clean = str(role_filter).lower().strip()
                
                for group_name in m['allowed_groups']:
                    g_clean = str(group_name).lower().strip()
                    if r_clean in g_clean or g_clean in r_clean:
                        is_primary = True
                        break
            
            if is_primary:
                results_prim.append(item)
            else:
                results_sec.append(item)
                
        return results_prim, results_sec

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter)

    pre_prim, pre_sec = [], []
    if view_mode != 'Operations':
        pre_prim, pre_sec = calculate_card_metrics(qs_pre, pre_metrics_db, 'pre')

    post_prim, post_sec = [], []
    post_prim, post_sec = calculate_card_metrics(qs_post, post_metrics_db, 'post')

    sbu_opts = sorted([s for s in Project.objects.values_list('sbu', flat=True).distinct() if s])
    
    context = {
        'view_mode': view_mode, 'start_date': start_str, 'end_date': end_str,
        'sbus': sbu_opts or ['North', 'South', 'West', 'Central'], 'selected_sbus': sbu_filter,
        'current_role': role_filter, 'people_opts': people_opts, 'selected_filters': selected_filters,
        'pre_prim': pre_prim, 'pre_sec': pre_sec, 'pre_count': pre_count,
        'post_prim': post_prim, 'post_sec': post_sec, 'post_count': post_count,
        'all_departments': all_departments,
    }
    return render(request, 'core/dashboard.html', context)

# ==============================================================================
# 3. EXPORT SUMMARY (Excel)
# ==============================================================================
def export_view(request):
    view_mode, _, _, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    projects = Project.objects.filter(sbu__in=sbu_filter)
    projects = _apply_people_filters(projects, view_mode, request)
    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)

    def generate_summary_df(queryset, metrics_list, prefix):
        total = queryset.count()
        data = [{"Metric Name": "TOTAL PROJECTS", "Threshold": "-", "Value": total, "%": "-"}]
        
        for m in metrics_list:
            param_name = f"thresh_{prefix}_{m['field']}"
            user_input = request.GET.get(param_name)
            try: threshold = float(user_input) if user_input else m['def']
            except: threshold = m['def']
            
            count = queryset.filter(**{f"{m['field']}__gte": threshold}).count()
            pct = round((count / total * 100), 1) if total > 0 else 0.0
            
            data.append({
                "Metric Name": m['label'], "Category": m['success_cat'],
                "Threshold": threshold, "Value": count, "%": f"{pct}%"
            })
        return pd.DataFrame(data)

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter)

    df_pre = generate_summary_df(qs_pre, pre_metrics_db, 'pre') if view_mode != 'Operations' else pd.DataFrame()
    df_post = generate_summary_df(qs_post, post_metrics_db, 'post')

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        if not df_pre.empty: df_pre.to_excel(writer, sheet_name='Pre-Stage', index=False)
        if not df_post.empty: df_post.to_excel(writer, sheet_name='Post-Stage', index=False)
        if df_pre.empty and df_post.empty: pd.DataFrame({'Info': ['No Data']}).to_excel(writer, sheet_name='Empty', index=False)

    buffer.seek(0)
    response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="Summary_{view_mode}.xlsx"'
    return response

# ==============================================================================
# 4. EXPORT DETAILED (Excel with ID/Roles)
# ==============================================================================
def export_detailed_view(request):
    view_mode, _, _, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    projects = Project.objects.filter(sbu__in=sbu_filter)
    projects = _apply_people_filters(projects, view_mode, request)
    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)

    def generate_detailed_df(queryset, metrics_list):
        if not metrics_list and not queryset.exists(): return pd.DataFrame()

        role_map = {
            'Sales': {'Sales Head': 'sales_head', 'Sales Lead': 'sales_lead'},
            'Design': {'DH': 'design_dh', 'DM': 'design_dm', 'ID': 'design_id', '3D': 'design_3d'},
            'Operations': {'BU Head': 'ops_head', 'SPM/PM': 'ops_pm', 'PM': 'ops_pm', 'SOM/OM': 'ops_om', 'OM': 'ops_om', 'SS': 'ops_ss', 'MEP': 'ops_mep', 'CSC': 'ops_csc'}
        }
        
        selected_role_cols = []
        if view_mode in role_map:
            if role_filter != 'All Roles':
                col = role_map[view_mode].get(role_filter)
                if col: selected_role_cols = [col]
            else:
                selected_role_cols = list(set(role_map[view_mode].values()))

        metric_fields = [m['field'] for m in metrics_list]
        fetch_fields = ['project_code', 'project_name'] + selected_role_cols + metric_fields
        
        data = list(queryset.values(*fetch_fields))
        if not data: return pd.DataFrame()

        df = pd.DataFrame(data)
        rename_map = {
            'project_code': 'Project ID', 'project_name': 'Project Name',
            'sales_head': 'Sales Head', 'sales_lead': 'Sales Lead',
            'design_dh': 'DH', 'design_dm': 'DM', 'design_id': 'ID', 'design_3d': '3D',
            'ops_head': 'Ops Head', 'ops_pm': 'PM', 'ops_om': 'OM', 'ops_ss': 'SS', 'ops_mep': 'MEP', 'ops_csc': 'CSC'
        }
        for m in metrics_list: rename_map[m['field']] = m['label']
        
        df = df.rename(columns=rename_map).fillna('')

        base_cols = ['Project ID', 'Project Name']
        role_headers = [rename_map.get(c, c) for c in selected_role_cols]
        metric_headers = [m['label'] for m in metrics_list]
        
        final_order = base_cols + role_headers + metric_headers
        final_order = [c for c in final_order if c in df.columns]
        
        df = df[final_order]

        sum_row = df.sum(numeric_only=True)
        sum_df = pd.DataFrame([sum_row], columns=df.columns)
        sum_df['Project Name'] = 'Grand Total'
        return pd.concat([df, sum_df], ignore_index=True).fillna('')

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter)

    df_pre = generate_detailed_df(qs_pre, pre_metrics_db) if view_mode != 'Operations' else pd.DataFrame()
    df_post = generate_detailed_df(qs_post, post_metrics_db)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        if not df_pre.empty: df_pre.to_excel(writer, sheet_name='Pre-Detailed', index=False)
        if not df_post.empty: df_post.to_excel(writer, sheet_name='Post-Detailed', index=False)
        if df_pre.empty and df_post.empty: pd.DataFrame({'Info': ['No Data']}).to_excel(writer, sheet_name='Empty', index=False)

    buffer.seek(0)
    response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="Detailed_{view_mode}.xlsx"'
    return response

# ==============================================================================
# 5. LIVE REPORT DETAILED (View)
# ==============================================================================
def report_detailed_view(request):
    # Reuse Logic from Export Detailed
    # NOTE: Since this is identical logic, we are re-using the exact same flow 
    # but rendering HTML instead of Excel.
    view_mode, start_str, end_str, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    projects = Project.objects.filter(sbu__in=sbu_filter)
    projects = _apply_people_filters(projects, view_mode, request)
    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)

    def generate_detailed_df(queryset, metrics_list):
        if not metrics_list and not queryset.exists(): return pd.DataFrame()
        
        # 1. Role Columns
        role_map = {
            'Sales': {'Sales Head': 'sales_head', 'Sales Lead': 'sales_lead'},
            'Design': {'DH': 'design_dh', 'DM': 'design_dm', 'ID': 'design_id', '3D': 'design_3d'},
            'Operations': {'BU Head': 'ops_head', 'SPM/PM': 'ops_pm', 'PM': 'ops_pm', 'SOM/OM': 'ops_om', 'OM': 'ops_om', 'SS': 'ops_ss', 'MEP': 'ops_mep', 'CSC': 'ops_csc'}
        }
        selected_role_cols = []
        if view_mode in role_map:
            if role_filter != 'All Roles':
                col = role_map[view_mode].get(role_filter)
                if col: selected_role_cols = [col]
            else:
                selected_role_cols = list(set(role_map[view_mode].values()))

        # 2. Data Fetch
        metric_fields = [m['field'] for m in metrics_list]
        fetch_fields = ['project_code', 'project_name'] + selected_role_cols + metric_fields
        data = list(queryset.values(*fetch_fields))
        if not data: return pd.DataFrame()

        df = pd.DataFrame(data)
        rename_map = {
            'project_code': 'Project ID', 'project_name': 'Project Name',
            'sales_head': 'Sales Head', 'sales_lead': 'Sales Lead',
            'design_dh': 'DH', 'design_dm': 'DM', 'design_id': 'ID', 'design_3d': '3D',
            'ops_head': 'Ops Head', 'ops_pm': 'PM', 'ops_om': 'OM', 'ops_ss': 'SS', 'ops_mep': 'MEP', 'ops_csc': 'CSC'
        }
        for m in metrics_list: rename_map[m['field']] = m['label']
        df = df.rename(columns=rename_map).fillna('')

        # 3. Order
        base_cols = ['Project ID', 'Project Name']
        role_headers = [rename_map.get(c, c) for c in selected_role_cols]
        metric_headers = [m['label'] for m in metrics_list]
        final_order = [c for c in (base_cols + role_headers + metric_headers) if c in df.columns]
        df = df[final_order]

        # 4. Total
        sum_row = df.sum(numeric_only=True)
        sum_df = pd.DataFrame([sum_row], columns=df.columns)
        sum_df['Project Name'] = 'Grand Total'
        return pd.concat([df, sum_df], ignore_index=True).fillna('')

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter)

    df_pre = generate_detailed_df(qs_pre, pre_metrics_db) if view_mode != 'Operations' else pd.DataFrame()
    df_post = generate_detailed_df(qs_post, post_metrics_db)

    context = {
        'view_mode': view_mode, 'start_date': start_str, 'end_date': end_str,
        'df_pre': df_pre.to_html(classes='table table-bordered table-striped table-hover table-sm', index=False, justify='center') if not df_pre.empty else None,
        'df_post': df_post.to_html(classes='table table-bordered table-striped table-hover table-sm', index=False, justify='center') if not df_post.empty else None,
    }
    return render(request, 'core/report_detailed.html', context)

# ==============================================================================
# 6. LEADERSHIP LIVE REPORT (Fixed - Renders Summary Tables)
# ==============================================================================
def report_view(request):
    """
    Renders the 'Leadership Summary' (Counts & Percentages) as a live HTML page.
    Reuses report_detailed.html but sends summary dataframes.
    """
    view_mode, start_str, end_str, start_dt, end_dt, sbu_filter, role_filter, roll_start, roll_end = _get_request_params(request)
    projects = Project.objects.filter(sbu__in=sbu_filter)
    projects = _apply_people_filters(projects, view_mode, request)
    qs_pre, qs_post = _get_stage_querysets(view_mode, projects, start_dt, end_dt, roll_start, roll_end)

    def generate_summary_df(queryset, metrics_list, prefix):
        total = queryset.count()
        data = [{"Metric Name": "TOTAL PROJECTS", "Threshold": "-", "Value": total, "%": "-"}]
        
        for m in metrics_list:
            param_name = f"thresh_{prefix}_{m['field']}"
            user_input = request.GET.get(param_name)
            try: threshold = float(user_input) if user_input else m['def']
            except: threshold = m['def']
            
            count = queryset.filter(**{f"{m['field']}__gte": threshold}).count()
            pct = round((count / total * 100), 1) if total > 0 else 0.0
            
            data.append({
                "Metric Name": m['label'], "Category": m['success_cat'],
                "Threshold": threshold, "Value": count, "%": f"{pct}%"
            })
        return pd.DataFrame(data)

    pre_metrics_db = _fetch_metrics_from_db(view_mode, 'Pre', role_filter)
    post_metrics_db = _fetch_metrics_from_db(view_mode, 'Post', role_filter)

    df_pre = generate_summary_df(qs_pre, pre_metrics_db, 'pre') if view_mode != 'Operations' else pd.DataFrame()
    df_post = generate_summary_df(qs_post, post_metrics_db, 'post')

    context = {
        'view_mode': view_mode, 'start_date': start_str, 'end_date': end_str,
        'df_pre': df_pre.to_html(classes='table table-bordered table-striped table-hover table-sm', index=False, justify='center') if not df_pre.empty else None,
        'df_post': df_post.to_html(classes='table table-bordered table-striped table-hover table-sm', index=False, justify='center') if not df_post.empty else None,
        'report_title': 'Leadership Summary Report' # Context flag for title
    }
    return render(request, 'core/report_detailed.html', context)



def project_detail(request, pk):
    # 1. Fetch Project or 404
    project = get_object_or_404(Project, pk=pk)
    
    # 2. Render Template
    return render(request, 'core/project_detail.html', {
        'project': project,
    })

# ==============================================================================
# 7. PROJECT SCORECARD (The Credit System)
# ==============================================================================
def project_scorecard_view(request, project_code):
    project = get_object_or_404(Project, project_code=project_code)
    
    # 1. Get role from URL
    raw_role_param = request.GET.get('metric_role', 'Design - ID')
    
    # 2. Extract search term (e.g., "ID" from "Design - ID")
    search_term = raw_role_param.split(' - ')[-1] if ' - ' in raw_role_param else raw_role_param
    
    # 3. Find UserGroup and handle "None" safety
    user_group = UserGroup.objects.filter(name__icontains=search_term).first()
    all_groups = UserGroup.objects.select_related('department').order_by('department__name', 'name')
    
    if not user_group:
         return render(request, 'core/project_scorecard.html', {
            'project': project, 
            'error': f"Role '{search_term}' not found.",
            'all_groups': all_groups,
            'selected_role_full': raw_role_param 
        })

    # 4. Stage & Metric Logic (Normalized)
    raw_stage = str(project.stage).strip().lower()
    metric_stage_key = 'Post' if any(x in raw_stage for x in ['post', 'exec', 'ops']) else 'Pre'
    
    metrics = Metric.objects.filter(stage=metric_stage_key, department=user_group.department).prefetch_related('metricweight_set')

    # 5. Calculation Engine
    total_factor_sum = 0
    temp_list = []
    for metric in metrics:
        weight_obj = metric.metricweight_set.filter(user_group=user_group).first()
        if weight_obj and weight_obj.factor > 0:
            total_factor_sum += weight_obj.factor
            current_value = getattr(project, metric.field_name, 0.0)
            threshold = metric.default_threshold
            raw_progress = (current_value / threshold) if threshold > 0 else (1.0 if current_value > 0 else 0.0)
            
            temp_list.append({
                'label': metric.label, 'factor': weight_obj.factor, 'raw_value': current_value,
                'threshold': threshold, 'capped_progress': min(raw_progress, 1.0),
                'display_percentage': int(raw_progress * 100)
            })

    final_scores = []
    earned_score_total = 0
    clean_fmt = lambda val: int(val) if val % 1 == 0 else round(val, 1)

    for item in temp_list:
        target_pct = (item['factor'] / total_factor_sum * 100) if total_factor_sum > 0 else 0
        points = item['capped_progress'] * target_pct
        earned_score_total += points
        final_scores.append({
            'metric': item['label'], 'factor': item['factor'], 'actual': clean_fmt(item['raw_value']),
            'target': clean_fmt(item['threshold']), 'target_percent': round(target_pct, 1),
            'status_text': f"{item['display_percentage']}%", 'points_earned': round(points, 1)
        })

    final_scores.sort(key=lambda x: x['factor'], reverse=True)

    return render(request, 'core/project_scorecard.html', {
        'project': project, 'user_group': user_group,
        'selected_group_id': user_group.id, # Fixed ID for UI selection
        'all_groups': all_groups, 'total_factor': total_factor_sum,
        'scores': final_scores, 'project_total': round(earned_score_total, 1),
        'project_total_int': int(round(earned_score_total, 0)),
        'metric_stage': metric_stage_key
    })

# ==============================================================================
# 8. LEADERBOARD VIEW (Gamification)
# ==============================================================================
def leaderboard_view(request):
    # 1. Get Params
    view_mode, start_str, end_str, start_dt, end_dt, _, _, _, _ = _get_request_params(request)
    
    # 2. Dynamic SBU Options
    all_sbu_options = list(Project.objects.exclude(sbu__isnull=True).exclude(sbu="").values_list('sbu', flat=True).distinct())
    all_sbu_options.sort()
    if not all_sbu_options: all_sbu_options = ['North', 'South', 'East', 'Central']

    if 'sbu' in request.GET: sbu_filter = request.GET.getlist('sbu')
    else: sbu_filter = all_sbu_options 

    # 3. Role Config
    selected_role_name = request.GET.get('role', 'Sales Lead')
    
    ROLE_CONFIG = {
        'Sales Lead':      {'field': 'sales_lead', 'url_param': 'f_s_lead', 'view': 'Sales'},
        'Sales Head':      {'field': 'sales_head', 'url_param': 'f_s_head', 'view': 'Sales'},
        'ID':              {'field': 'design_id',  'url_param': 'f_d_id',   'view': 'Design'},
        '3D':              {'field': 'design_3d',  'url_param': 'f_d_3d',   'view': 'Design'},
        'DM':              {'field': 'design_dm',  'url_param': 'f_d_dm',   'view': 'Design'},
        'DH':              {'field': 'design_dh',  'url_param': 'f_d_dh',   'view': 'Design'},
        'Cluster/BU Head': {'field': 'ops_head',   'url_param': 'f_o_head', 'view': 'Operations'},
        'SPM/PM':          {'field': 'ops_pm',     'url_param': 'f_o_pm',   'view': 'Operations'},
        'SOM/OM':          {'field': 'ops_om',     'url_param': 'f_o_om',   'view': 'Operations'},
        'SS':              {'field': 'ops_ss',     'url_param': 'f_o_ss',   'view': 'Operations'},
        'MEP':             {'field': 'ops_mep',    'url_param': 'f_o_mep',  'view': 'Operations'},
        'CSC':             {'field': 'ops_csc',    'url_param': 'f_o_csc',  'view': 'Operations'},
    }

    simple_role_name = selected_role_name.split(' - ')[-1] if ' - ' in selected_role_name else selected_role_name
    config = ROLE_CONFIG.get(simple_role_name)
    
    if not config:
        return render(request, 'core/leaderboard.html', {'error': f"Role '{selected_role_name}' not supported.", 'all_roles': sorted(ROLE_CONFIG.keys())})

    project_field = config['field']
    user_group = UserGroup.objects.filter(name__icontains=simple_role_name).first()
    if not user_group:
         return render(request, 'core/leaderboard.html', {'error': f"User Group '{selected_role_name}' not found.", 'all_roles': sorted(ROLE_CONFIG.keys())})

    # 5. Fetch Projects (Apply Filters & Exclude Test)
    projects = Project.objects.filter(sbu__in=sbu_filter)\
        .exclude(**{f"{project_field}__isnull": True})\
        .exclude(**{f"{project_field}__exact": ""})\
        .exclude(project_code__isnull=True)\
        .exclude(project_code__exact="")\
        .exclude(project_code="PS-02AUG23-BB1_TEST-SOMERSET-01")

    projects = projects.filter(Q(login_date__range=[start_dt, end_dt]) | Q(start_date__range=[start_dt, end_dt])).distinct()

    # 6. Scoring Engine
    metrics = Metric.objects.filter(metricweight__user_group=user_group, metricweight__factor__gt=0).distinct().prefetch_related('metricweight_set')
    stage_totals = {}
    valid_metrics = []

    for m in metrics:
        weight_obj = m.metricweight_set.filter(user_group=user_group).first()
        if weight_obj and weight_obj.factor > 0:
            stage_totals[m.stage] = stage_totals.get(m.stage, 0) + weight_obj.factor
            valid_metrics.append({'field': m.field_name, 'factor': weight_obj.factor, 'stage': m.stage, 'threshold': m.default_threshold})

    # 7. Calculate Scores
    leaderboard = {}
    for proj in projects:
        if not proj.project_code or not str(proj.project_code).strip():
            continue

        user_email = getattr(proj, project_field)
        if not user_email: continue
        
        user_key = str(user_email).strip().lower()
        if user_key not in leaderboard:
            leaderboard[user_key] = {'name': user_email, 'total_score': 0, 'projects': 0, 'breakdown': []}

        raw_stage = str(proj.stage).strip().lower()
        current_stage = 'Post' if any(x in raw_stage for x in ['post', 'exec', 'ops', 'handover']) else 'Pre'
        total_possible = stage_totals.get(current_stage, 0)
        
        project_score = 0
        if total_possible > 0:
            for vm in valid_metrics:
                if vm['stage'] == current_stage:
                    val = getattr(proj, vm['field'], 0.0)
                    threshold = vm['threshold']
                    
                    if threshold > 0: raw_progress = val / threshold
                    else: raw_progress = 1.0 if val > 0 else 0.0
                    
                    capped_progress = min(raw_progress, 1.0)
                    project_score += capped_progress * (vm['factor'] / total_possible * 100)
        
        project_score = round(project_score, 1)

        leaderboard[user_key]['total_score'] += project_score
        leaderboard[user_key]['projects'] += 1
        leaderboard[user_key]['breakdown'].append({
            'project_name': proj.project_name or proj.project_code,
            'code': proj.project_code,
            'stage': current_stage,
            'sbu': proj.sbu,
            'score': project_score
        })

    # 8. Post-Processing
    sorted_leaderboard = sorted(leaderboard.values(), key=lambda x: x['total_score'], reverse=True)
    all_scores = [u['total_score'] for u in leaderboard.values()]
    total_users = len(all_scores)

    for idx, row in enumerate(sorted_leaderboard, 1):
        row['rank'] = idx
        row['breakdown'].sort(key=lambda x: x['score'], reverse=True)

        if total_users > 0:
            people_beaten = sum(1 for s in all_scores if s <= row['total_score'])
            row['percentile'] = int((people_beaten / total_users) * 100)
        else:
            row['percentile'] = 0
            
        row['total_score'] = round(row['total_score'], 1)

    context = {
        'leaderboard': sorted_leaderboard,
        'selected_role': selected_role_name,
        'all_roles': sorted(ROLE_CONFIG.keys()),
        'start_date': start_str,
        'end_date': end_str,
        'selected_sbus': sbu_filter, 
        'sbu_options': all_sbu_options,
        'link_view': config.get('view', 'Sales'),
        'link_param': config.get('url_param', 'f_s_lead')
    }
    return render(request, 'core/leaderboard.html', context)