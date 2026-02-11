# Keeps METRICS_CONFIG & EXCEL_COL_MAP  (CHANGES NEED TO MADE FOR COL NAMES)
DEPT_CHOICES = (
    ('Marketing', 'Marketing'),
    ('Sales', 'Sales'),
    ('Design', 'Design'),
    ('Operations', 'Operations'),
    ('Purchase', 'Purchase'),
    ('Finance', 'Finance'),
)

DEPT_PEOPLE_MAP = {
    'Marketing': [
        ('m_head', 'Marketing Head'),
        ('m_lead', 'Marketing Lead'),
    ],
    'Sales': [
        ('s_head', 'Sales Head'),
        ('s_lead', 'Sales Lead'),
    ],
    'Design': [
        ('d_dh', 'DH'),
        ('d_dm', 'DM'),
        ('d_id', 'ID'),
        ('d_3d', '3D'),
    ],
    'Operations': [
        ('o_head', 'Cluster/BU Head'),
        ('o_pm', 'SPM/PM'),
        ('o_om', 'SOM/OM'),
        ('o_ss', 'SS'),
        ('o_mep', 'MEP'),
        ('o_csc', 'CSC'),
    ],
    'Purchase': [
        ('p_head', 'Purchase Head'),
        ('p_mgr', 'Purchase Manager'),
        ('p_exec', 'Purchase Executive'),
    ],
    'Finance': [
        ('f_head', 'Finance Head'), 
    ],
}

METRICS_CONFIG = {
    "Sales": {
        "Pre": [
            {"label": "Requirements Uploaded", "field": "req_uploaded", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
            {"label": "Site Visit Reports", "field": "site_visit_report", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
            {"label": "Client Visits", "field": "client_access", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
        ],
        "Post": [
            {"label": "Requirements Uploaded", "field": "req_uploaded", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
            {"label": "Site Visit Reports", "field": "site_visit_report", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
            {"label": "Client Visits", "field": "client_access", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
            {"label": "BOQs Uploaded", "field": "boq_uploaded", "def": 1.0, "roles": ["Sales Lead", "Sales Head"]},
            {"label": "Contracts Uploaded", "field": "contract_uploaded", "def": 1.0, "roles": ["Sales Head"]},
        ]
    },
    "Design": {
        "Pre": [
            # ID
            {"label": "Pins Mapped", "field": "mapped_spaces", "def": 3.0, "roles": ["ID"]},
            {"label": "Key Plans to Spaces", "field": "key_plans_ratio", "def": 0.33, "roles": ["ID"]},
            {"label": "No Key Plans Spaces", "field": "no_plans_for_key_spaces", "def": 1.0, "roles": ["ID"]},
            
            # 3D
            {"label": "Renders", "field": "renders", "def": 5.0, "roles": ["3D", "DM", "DH"]}, # User said 3D=Upload Renders
            
            # DM/DH
            {"label": "Furniture Layouts", "field": "furniture_layouts", "def": 1.0, "roles": ["DM", "DH"]},
            {"label": "Approved Furniture Layouts", "field": "approved_layouts", "def": 1.0, "roles": ["DM", "DH"]},
            {"label": "Approved Renders", "field": "approved_renders", "def": 5.0, "roles": ["DM", "DH"]},
            {"label": "Material Deck", "field": "material_deck", "def": 1.0, "roles": ["DM", "DH"]},
        ],
        "Post": [
            # ID
            {"label": "Other Layouts", "field": "other_layouts", "def": 1.0, "roles": ["ID"]},
            {"label": "Furniture Layouts", "field": "furniture_layouts", "def": 1.0, "roles": ["ID", "DM", "DH"]}, # Shared
            {"label": "Mapped Spaces", "field": "mapped_spaces", "def": 3.0, "roles": ["ID"]},
            {"label": "No Key Plans Spaces", "field": "no_plans_for_key_spaces", "def": 1.0, "roles": ["ID"]},
            {"label": "TD & Elevations", "field": "td_elevations", "def": 10.0, "roles": ["ID"]},
            {"label": "CAD Files", "field": "cad_files", "def": 10.0, "roles": ["ID"]},
            {"label": "GFC Download", "field": "gfc_download", "def": 7.0, "roles": ["ID", "DM", "DH"]}, # Shared

            # 3D
            {"label": "Renders", "field": "renders", "def": 5.0, "roles": ["3D", "DM", "DH"]}, 
            {"label": "Approved Renders", "field": "approved_renders", "def": 5.0, "roles": ["3D", "DM", "DH"]}, 

            # DM/DH (Specifics)
            {"label": "Approved Furniture Layouts", "field": "approved_layouts", "def": 1.0, "roles": ["DM", "DH"]},
            {"label": "Slides Downloaded", "field": "slides_download", "def": 1.0, "roles": ["DM", "DH"]},
            {"label": "Material Deck", "field": "material_deck", "def": 1.0, "roles": ["DM", "DH"]},
            {"label": "Client Visits", "field": "client_visit_des", "def": 7.0, "roles": ["DM", "DH"]},
        ]
    },
    "Operations": [
        # MEP
        {"label": "MEP Drawings", "field": "mep_drawings", "def": 3.0, "roles": ["BU Head", "MEP"]},

        # SPM / PM
        {"label": "GRNs/SRNs Approved", "field": "grn_approved", "def": 3.0, "roles": ["BU Head", "SPM", "PM"]},
        {"label": "WPRs Downloaded", "field": "wpr_download", "def": 3.0, "roles": ["BU Head", "SPM", "PM"]},
        {"label": "WPRs Shared", "field": "wpr_shared", "def": 3.0, "roles": ["BU Head", "SPM", "PM"]},
        {"label": "Invoices / Receipts", "field": "invoices", "def": 1.0, "roles": ["BU Head", "SPM", "PM"]},

        # SS
        {"label": "DPR Added (% days)", "field": "dpr_ratio", "def": 0.5, "roles": ["BU Head", "SS"]},
        {"label": "GRNs/SRNs Created", "field": "grn_created", "def": 3.0, "roles": ["BU Head", "SS"]},

        # General / Unassigned (Will show in All Roles, but hidden for specific roles unless added)
        {"label": "Site Images", "field": "site_images", "def": 30.0, "roles": ["BU Head", "SS", "PM"]},
        {"label": "Handover Documents", "field": "handover_docs", "def": 3.0, "roles": ["BU Head", "PM"]},
        {"label": "Unique Weekly Tasks", "field": "weekly_tasks", "def": 10.0, "roles": ["BU Head", "PM"]},
        {"label": "Unique Daily Tasks", "field": "daily_tasks", "def": 30.0, "roles": ["BU Head", "SS", "PM"]},
        {"label": "Manpower Ratio", "field": "manpower_ratio", "def": 0.66, "roles": ["BU Head", "SS", "PM"]},
    ],

    "Marketing": [],

    "Fiannce": [],

    "Purchase": [],
    
}

# Used for mapping Excel Column names to Django Model Fields
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
    "WPRs Download": "wpr_download",
    "WPR Shared Client": "wpr_shared",
    "Total Unique Weekly Tasks": "weekly_tasks",
    "Total Unique Daily Tasks": "daily_tasks",
    "Total GRN/SRN": "grn_created",
    "Total Approved GRN/SRN": "grn_approved"
}