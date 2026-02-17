# core/constants.py

# ==============================================================================
# DEPARTMENT CONFIGURATION
# ==============================================================================
DEPT_CHOICES = (
    ('Marketing', 'Marketing'),
    ('Sales', 'Sales'),
    ('Design', 'Design'),
    ('Operations', 'Operations'),
    ('Purchase', 'Purchase'),
    ('Finance', 'Finance'),
)

# Maps Department View Mode to specific Database Fields for filtering.
# CRITICAL FIX: The first item in the tuple MUST be the exact field name in models.py
DEPT_PEOPLE_MAP = {
    'Marketing': [
        # Marketing fields don't exist in your Model yet, so we leave this empty 
        # to prevent crashes. Uncomment/Update if you add 'marketing_head' to models.py
        # ('marketing_head', 'Marketing Head'), 
    ],
    'Sales': [
        ('sales_head', 'Sales Head'), 
        ('sales_lead', 'Sales Lead')
    ],
    'Design': [
        ('design_dh', 'DH'), 
        ('design_dm', 'DM'), 
        ('design_id', 'ID'), 
        ('design_3d', '3D')
    ],
    'Operations': [
        ('ops_head', 'Cluster/BU Head'), 
        ('ops_pm', 'SPM/PM'),
        ('ops_om', 'SOM/OM'), 
        ('ops_ss', 'SS'),
        ('ops_mep', 'MEP'), 
        ('ops_csc', 'CSC')
    ],
    'Purchase': [
        # Purchase fields don't exist in Model yet
        # ('purchase_head', 'Purchase Head'), 
    ],
    'Finance': [
        # Finance fields don't exist in Model yet
        # ('finance_head', 'Finance Head'), 
    ],
}

# ==============================================================================
# ROLE DEFINITIONS (Used in Leaderboards & Links)
# ==============================================================================
# Format: 'Display Name': {'field': 'db_column_name', 'link': 'url_param', 'dept': 'Department'}
ROLE_CONFIG = {
    # Sales
    'Sales Lead':      {'field': 'sales_lead', 'link': 'f_s_lead', 'dept': 'Sales'},
    'Sales Head':      {'field': 'sales_head', 'link': 'f_s_head', 'dept': 'Sales'},
    
    # Design
    'ID':              {'field': 'design_id',  'link': 'f_d_id',   'dept': 'Design'},
    '3D':              {'field': 'design_3d',  'link': 'f_d_3d',   'dept': 'Design'},
    'DM':              {'field': 'design_dm',  'link': 'f_d_dm',   'dept': 'Design'},
    'DH':              {'field': 'design_dh',  'link': 'f_d_dh',   'dept': 'Design'},
    
    # Operations (Standardized Keys)
    'Cluster/BU Head': {'field': 'ops_head',   'link': 'f_o_head', 'dept': 'Operations'},
    'SPM/PM':          {'field': 'ops_pm',     'link': 'f_o_pm',   'dept': 'Operations'},
    'SOM/OM':          {'field': 'ops_om',     'link': 'f_o_om',   'dept': 'Operations'},
    'SS':              {'field': 'ops_ss',     'link': 'f_o_ss',   'dept': 'Operations'},
    'MEP':             {'field': 'ops_mep',    'link': 'f_o_mep',  'dept': 'Operations'},
    'CSC':             {'field': 'ops_csc',    'link': 'f_o_csc',  'dept': 'Operations'},

    # Purchase
    'Purchase Head':      {'field': 'p_head', 'link': 'f_p_head', 'dept': 'Purchase'},
    'Purchase Manager':   {'field': 'p_mgr',  'link': 'f_p_mgr',  'dept': 'Purchase'},
    'Purchase Executive': {'field': 'p_exec', 'link': 'f_p_exec', 'dept': 'Purchase'},

    # Finance
    'Finance Head':       {'field': 'f_head', 'link': 'f_f_head', 'dept': 'Finance'},
    
    # Marketing
    'Marketing Head':     {'field': 'm_head', 'link': 'f_m_head', 'dept': 'Marketing'},
    'Marketing Lead':     {'field': 'm_lead', 'link': 'f_m_lead', 'dept': 'Marketing'},
}

# ==============================================================================
# EXCEL MAPPING & METRICS DEFAULTS
# ==============================================================================
# Used to map Excel Column names to Django Model Fields
EXCEL_COL_MAP = {
    # Sales
    "Requirements": "req_uploaded",
    "Site Visit Report": "site_visit_report",
    "Client Access": "client_access",
    "BOQ": "boq_uploaded",
    "Contract": "contract_uploaded",
    
    # Design
    "Furniture Layouts": "furniture_layouts",
    "Approved Furniture Layouts": "approved_layouts",
    "Mapped Spaces": "mapped_spaces",
    "No Key Plans Spaces": "no_plans_for_key_spaces",
    "Renders": "renders",
    "Approved Renders": "approved_renders",
    "TD & Elevations": "td_elevations",
    "CAD Files": "cad_files",
    "Design Slides Download": "slides_download",
    "Material Deck Download": "material_deck",
    "GFC Download": "gfc_download",
    "Client Visit": "client_visit_des",
    
    # Ops
    "Site Progress Images": "site_images",
    "Invoices, Receipts": "invoices",
    "MEP Drawings": "mep_drawings",
    "Handover Documents": "handover_docs",
    "WPR Download": "wpr_download",
    "WPR Share to Client": "wpr_shared",
    "Total Unique Weekly Tasks": "weekly_tasks",
    "Total Unique Daily Tasks": "daily_tasks",
    "Total GRN/SRN": "grn_created",
    "Total Approved GRN/SRN": "grn_approved",
    "Weeks Till Date": "weeks_till_date",
    "Days Till Date": "days_till_date",
    "WPR Download Weeks": "wpr_download_weeks",
    "Manpower Added Days": "manpower_added_days",

}

# ==============================================================================
# REPORT COLUMN ORDERING 
# ==============================================================================
# Use this to define the order of columns in the report_detailed view
# These names must match the 'Metric Label' in Django Admin EXACTLY.

COMMON_REPORT_COLS = [
    'Project Name', 
    'Project Code', 
    'Lead ID', 
    'SBU', 
    'Stage', 
    'Project Type', 
    'Floors'
]

REPORT_ORDER_CONFIG = {
    'Sales': {   
        'Pre': COMMON_REPORT_COLS + [
            'Requirements Uploaded',  
            'Site Visit Reports',  
            'Client Visits',              
            'Sales Head',          
            'Sales Lead'          
        ],
        'Post': COMMON_REPORT_COLS + [
            'BOQs Uploaded',
            'Contracts Uploaded',
            'Requirements Uploaded',  
            'Site Visit Reports',  
            'Client Visits',              
            'Sales Head',          
            'Sales Lead'           
        ]
    },

    'Design': {
        'Pre': COMMON_REPORT_COLS + [
            'Layouts',
            'Furniture Layouts',
            'Approved Furniture Layouts',
            'Other Layouts',
            'Pins Mapped',
            'No Key Plans Spaces',
            'Key Plans to Spaces Ratio',
            'Renders',                                
            'Approved Renders',                                                                
            'Material Deck', 
            'Client Visits',                                                               
            'Design Head (DH)',         
            'Design Manager (DM)',         
            'Interior Designer (ID)',   
            '3D Visualizer (3D)'        
        ],
        'Post': COMMON_REPORT_COLS + [
            'Layouts',
            'Furniture Layouts',
            'Approved Furniture Layouts',
            'Other Layouts',
            'Mapped Spaces',
            'No Key Plans Spaces',
            'Key Plans to Spaces Ratio',
            'Renders',                                
            'Approved Renders',                                                                
            'Material Deck',
            'Client Visits',   
            'GFC Download',
            'Slides Downloaded',
            'TD & Elevations',
            'CAD Files',
            'Design Head (DH)',         
            'Design Manager (DM)',         
            'Interior Designer (ID)',   
            '3D Visualizer (3D)'       
        ]

    },

    'Operations': COMMON_REPORT_COLS + [
        'Site Images',                 
        'Invoices / Receipts',                 
        'MEP Drawings',                 
        'Handover Documents',                 
        'Weeks Till Date',                          
        'WPR Download Weeks',                     
        'WPR Download',                     
        'WPR Shared',   
        'Unique Weekly Tasks',              
        'Unique Daily Tasks',              
        'Unique Daily Tasks',              
        'Manpower Ratio',
        'Days Till Date',
        'DPR Added (% days)',
        'Manpower Added Days',
        'GRNs/SRNs Created',                     
        'GRNs/SRNs Approved',                     
        'Cluster/BU Head',            
        'SPM/PM',     
        'SOM/OM',     
        'SS',
        'MEP',
        'CSC'      
    ],

    'Purchase': COMMON_REPORT_COLS + [
        'Purchase Head',
        'Purchase Manager'
    ],

    'Marketing': COMMON_REPORT_COLS,
    'Finance': COMMON_REPORT_COLS,
}