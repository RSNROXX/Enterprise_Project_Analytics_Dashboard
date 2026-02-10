import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from django.shortcuts import render, redirect, get_object_or_404  #type:ignore
from django.contrib import messages                             #type:ignore
from django.http import HttpResponse                            #type:ignore
from django.db.models import Q                                  #type:ignore

from .forms import UploadFileForm
from .models import Project, Metric, Department, UserGroup 
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
    metrics_qs = Metric.objects.filter(department=dept, stage=stage).prefetch_related('visible_to_groups')

    metrics_list = []
    for m in metrics_qs:
        # Store the allowed groups so we can sort them later in Python
        allowed_groups = [g.name for g in m.visible_to_groups.all()]
        
        metrics_list.append({
            'label': m.label,
            'field': m.field_name,
            'def': m.default_threshold,
            'success_cat': m.success_metric.name if m.success_metric else None,
            'success_color': m.success_metric.color if m.success_metric else 'secondary', 
            'allowed_groups': allowed_groups, # New: We need this for sorting
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
                
                for name in xls.sheet_names:
                    lower = str(name).lower()
                    if 'sales' in lower: sheet_map['sales'] = str(name)
                    elif 'design' in lower: sheet_map['design'] = str(name)
                    elif 'operation' in lower or 'ops' in lower: sheet_map['operation'] = str(name)

                project_data_map = {}

                def clean_str(val): return '' if pd.isnull(val) or str(val).strip().lower() == 'nan' else str(val).strip()
                def parse_date(val): return pd.to_datetime(val).date() if pd.notnull(val) else None
                def clean_num(val):
                    if pd.isnull(val): return 0.0
                    try: return float(str(val).replace('%','').replace(',','').strip())
                    except: return 0.0
                
                def get_clean_id(row):
                    for c in ['Project Code','project_code','Code','code','Lead Id']:
                        if c in row.index and pd.notnull(row[c]):
                            v = str(row[c]).strip().upper()
                            if v and v != 'NAN': return v.replace('.0','')
                    return None

                def process_sheet(sheet_name, sheet_type):
                    if not sheet_name: return
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    df.columns = [str(c).strip() for c in df.columns]

                    if sheet_type == 'design':
                        if 'No Key Plans Spaces' in df.columns and 'Mapped Spaces' in df.columns:
                            df['Key Plans Ratio'] = pd.to_numeric(df['No Key Plans Spaces'], errors='coerce').fillna(0) / pd.to_numeric(df['Mapped Spaces'], errors='coerce').replace(0,1).fillna(0)
                        if 'Layouts' in df.columns and 'Furniture Layouts' in df.columns:
                            df['Other Layouts'] = df['Layouts'] - df['Furniture Layouts']
                    
                    if sheet_type == 'operation':
                        ops_calcs = [('WPR Half Week','WPR Download Weeks','Weeks Till Date'), ('Manpower Ratio','Actual Manpower','Planned Manpower'),
                                     ('DPR Ratio','DPR Added Days','Days Till Date'), ('Manpower Day Ratio','Manpower Added Days','Days Till Date')]
                        for t, n, d in ops_calcs:
                            if n in df.columns and d in df.columns:
                                df[t] = (pd.to_numeric(df[n], errors='coerce') / pd.to_numeric(df[d], errors='coerce').replace(0,1)).fillna(0)

                    for _, row in df.iterrows():
                        p_id = get_clean_id(row)
                        if not p_id: continue
                        if p_id not in project_data_map: project_data_map[p_id] = {'project_code': p_id}
                        
                        meta = {
                            'project_name': ['Project Name'], 'sbu': ['SBU'], 'stage': ['Stage', 'Status'],
                            'sales_head': ['Sales Head'], 'sales_lead': ['Sales Lead'],
                            'design_dh': ['Design Head', 'DH'], 'design_dm': ['Design Lead', 'DM'], 'design_id': ['Design ID', 'ID'], 'design_3d': ['3D Visualizer', '3D'],
                            'ops_head': ['Ops Head'], 'ops_pm': ['Project Manager', 'SPM/PM'], 'ops_om': ['Ops Manager', 'SOM/OM'], 'ops_ss': ['Site Supervisor', 'SS'],
                            'ops_mep': ['MEP'], 'ops_csc': ['CSC']
                        }
                        for db, opts in meta.items():
                            for col in opts:
                                if col in row.index: 
                                    v = clean_str(row[col])
                                    if v: project_data_map[p_id][db] = v
                                    break
                        
                        dates = {'login_date': ['Project Login Date'], 'start_date': ['Project Start Date'], 'end_date': ['Project End Date']}
                        for db, opts in dates.items():
                            for col in opts:
                                if col in row.index: 
                                    v = parse_date(row[col])
                                    if v: project_data_map[p_id][db] = v
                                    break

                        full_map = EXCEL_COL_MAP.copy()
                        full_map.update({
                            'Key Plans Ratio':'key_plans_ratio', 'Other Layouts':'other_layouts', 
                            'WPR Half Week':'wpr_half_week', 'Manpower Ratio':'manpower_ratio', 
                            'DPR Ratio':'dpr_ratio', 'Manpower Day Ratio':'manpower_day_ratio'
                        })
                        
                        for xl, db in full_map.items():
                            if xl in row.index:
                                v = clean_num(row[xl])
                                if v != 0 or db not in project_data_map[p_id]: 
                                    project_data_map[p_id][db] = v

                process_sheet(sheet_map['sales'], 'sales')
                process_sheet(sheet_map['design'], 'design')
                process_sheet(sheet_map['operation'], 'operation')

                Project.objects.all().delete()
                Project.objects.bulk_create([Project(**d) for d in project_data_map.values()])
                
                messages.success(request, f"Successfully processed {len(project_data_map)} projects.")
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

            # --- LOGIC RESTORED: Primary vs Secondary ---
            # If "All Roles" -> Show Everything in Primary
            # If Role Selected -> Show Matching in Primary, Others in Secondary ("See More")
            is_primary = False
            if role_filter == "All Roles":
                is_primary = True
            elif role_filter in m['allowed_groups']:
                is_primary = True
            
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