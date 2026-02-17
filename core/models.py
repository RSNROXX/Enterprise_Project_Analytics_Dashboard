from django.db import models            # type: ignore

# ==============================================================================
# 1. CORE DATA MODEL (The Project)
# ==============================================================================

class Project(models.Model):
    """
    Represents a single project entry imported from the Excel sheet.
    Acts as the central entity for all scoring and reporting.
    """
    # --- Identification ---
    project_code = models.CharField(max_length=50, unique=True, help_text="Unique Identifier (e.g. FS-MUM-001)")
    project_name = models.CharField(max_length=255, null=True, blank=True)
    sbu = models.CharField(max_length=50, null=True, blank=True, verbose_name="Region/SBU")
    stage = models.CharField(max_length=50, null=True, blank=True, verbose_name="Project Stage")
    floors = models.CharField(max_length=50, null=True, blank=True)
    project_type = models.CharField(max_length=100, null=True, blank=True)
    lead_id = models.CharField(max_length=100, null=True, blank=True)
    
    # --- Key Dates ---
    login_date = models.DateField(null=True, blank=True)
    start_date = models.DateField(null=True, blank=True)
    end_date = models.DateField(null=True, blank=True)
    
    # --- Stakeholders (Team) ---
    sales_head = models.CharField(max_length=100, null=True, blank=True, verbose_name="Sales Head")
    sales_lead = models.CharField(max_length=100, null=True, blank=True, verbose_name="Sales Lead")

    design_dh = models.CharField(max_length=100, null=True, blank=True, verbose_name="DH")
    design_dm = models.CharField(max_length=100, null=True, blank=True, verbose_name="DM")
    design_id = models.CharField(max_length=100, null=True, blank=True, verbose_name="ID")
    design_3d = models.CharField(max_length=100, null=True, blank=True, verbose_name="3D")

    ops_head = models.CharField(max_length=100, null=True, blank=True, verbose_name="Cluster/BU Head")
    ops_pm = models.CharField(max_length=100, null=True, blank=True, verbose_name="SPM/PM")
    ops_om = models.CharField(max_length=100, null=True, blank=True, verbose_name="SOM/OM")
    ops_ss = models.CharField(max_length=100, null=True, blank=True, verbose_name="SS")
    ops_mep = models.CharField(max_length=100, null=True, blank=True, verbose_name="MEP")
    ops_csc = models.CharField(max_length=100, null=True, blank=True, verbose_name="CSC")

    m_head = models.CharField(max_length=100, null=True, blank=True, verbose_name="Marketing Head")
    m_lead = models.CharField(max_length=100, null=True, blank=True, verbose_name="Marketing Lead")
    
    p_head = models.CharField(max_length=100, null=True, blank=True, verbose_name="Purchase Head")
    p_mgr = models.CharField(max_length=100, null=True, blank=True, verbose_name="Purchase Manager")
    p_exec = models.CharField(max_length=100, null=True, blank=True, verbose_name="Purchase Executive")
    
    f_head = models.CharField(max_length=100, null=True, blank=True, verbose_name="Finance Head")

    # =========================================================
    # RAW METRICS (Populated via Excel Upload)
    # =========================================================
    
    # Sales
    req_uploaded = models.FloatField(default=0.0)
    site_visit_report = models.FloatField(default=0.0)
    client_access = models.FloatField(default=0.0)
    boq_uploaded = models.FloatField(default=0.0)
    contract_uploaded = models.FloatField(default=0.0)
    boq = models.FloatField(default=0.0)         
    contract = models.FloatField(default=0.0)    

    # Design
    furniture_layouts = models.FloatField(default=0.0)
    approved_layouts = models.FloatField(default=0.0)
    mapped_spaces = models.FloatField(default=0.0)
    no_plans_for_key_spaces = models.FloatField(default=0.0)
    renders = models.FloatField(default=0.0)
    approved_renders = models.FloatField(default=0.0)
    td_elevations = models.FloatField(default=0.0)
    cad_files = models.FloatField(default=0.0)
    slides_download = models.FloatField(default=0.0)
    material_deck = models.FloatField(default=0.0)
    gfc_download = models.FloatField(default=0.0)
    client_visit_des = models.FloatField(default=0.0)
    
    # Calculated Design Metrics
    key_plans_ratio = models.FloatField(default=0.0)
    other_layouts = models.FloatField(default=0.0)

    # Operations
    site_images = models.FloatField(default=0.0)
    invoices = models.FloatField(default=0.0)
    mep_drawings = models.FloatField(default=0.0)
    handover_docs = models.FloatField(default=0.0)
    wpr_download = models.FloatField(default=0.0)
    wpr_shared = models.FloatField(default=0.0)
    weekly_tasks = models.FloatField(default=0.0)
    daily_tasks = models.FloatField(default=0.0)
    grn_created = models.FloatField(default=0.0)
    grn_approved = models.FloatField(default=0.0)
    weeks_till_date = models.FloatField(default=0.0)
    days_till_date  = models.FloatField(default=0.0)
    wpr_download_weeks  = models.FloatField(default=0.0)
    manpower_added_days = models.FloatField(default=0.0)
    
    # Calculated Ops Metrics
    wpr_half_week = models.FloatField(default=0.0)
    manpower_ratio = models.FloatField(default=0.0)
    dpr_ratio = models.FloatField(default=0.0)
    wpr_ratio = models.FloatField(default=0.0)
    manpower_day_ratio = models.FloatField(default=0.0)

    def __str__(self):
        return f"{self.project_code} - {self.project_name}"


# ==============================================================================
# 2. CONFIGURATION MODELS (Admin Managed)
# ==============================================================================

class Department(models.Model):
    """ High-level functional areas (e.g. Sales, Design). """
    name = models.CharField(max_length=50, unique=True)

    def __str__(self):
        return self.name

class UserGroup(models.Model):
    """ Specific Roles within a Department (e.g. 'Sales Lead', '3D Visualizer'). """
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    name = models.CharField(max_length=50)

    def __str__(self):
        return f"{self.department.name} - {self.name}"
    
class SuccessMetric(models.Model):
    COLOR_CHOICES = [
        ('primary', 'Blue (Primary)'),
        ('secondary', 'Grey (Secondary)'),
        ('success', 'Green (Success)'),
        ('danger', 'Red (Danger)'),
        ('warning', 'Yellow (Warning)'),
        ('info', 'Cyan (Info)'),
        ('light', 'White (Light)'),
        ('dark', 'Black (Dark)'),
    ]

    name = models.CharField(max_length=50, unique=True)
    color = models.CharField(max_length=20, choices=COLOR_CHOICES, default='secondary')

    def __str__(self):
        return self.name

class Metric(models.Model):
    """ 
        Defines a KPI to be tracked.
        Maps a user-friendly Label (e.g. 'Client Visits') to a Database Field (e.g. 'client_access').
    """
    STAGE_CHOICES = [('Pre', 'Pre-Stage'), ('Post', 'Post-Stage')]

    label = models.CharField(max_length=100) 
    field_name = models.CharField(max_length=100, help_text="Must match a field in the Project model exactly.")
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    stage = models.CharField(max_length=10, choices=STAGE_CHOICES)

    # Min and Max Thresholds 
    min_threshold = models.FloatField(default=1.0, help_text="Below this value, score is 0.")
    max_threshold = models.FloatField(default=10.0, help_text="Maximum points achievable for this metric.")
    
    # Logic
    success_metric = models.ForeignKey(SuccessMetric, on_delete=models.SET_NULL, null=True, blank=True)
    
    # Legacy Visibility (kept for backward compatibility, but MetricWeight is preferred)
    visible_to_groups = models.ManyToManyField(UserGroup, blank=True)

    def __str__(self):
        return f"{self.label} ({self.department})"

class MetricWeight(models.Model):
    """
        The Core Scoring Engine Configuration.
        Assigns an 'Importance Factor' to a Metric for a specific User Group.
    """
    metric = models.ForeignKey(Metric, on_delete=models.CASCADE)
    user_group = models.ForeignKey(UserGroup, on_delete=models.CASCADE)
    
    WEIGHT_CHOICES = [
        (1, '1 - Minimal'), (2, '2 - Very Low'), (3, '3 - Low'),
        (4, '4 - Low-Medium'), (5, '5 - Medium (Standard)'),
        (6, '6 - Medium-High'), (7, '7 - High'), (8, '8 - Very High'),
        (9, '9 - Critical'), (10, '10 - Maximum Priority'),
    ]
    
    factor = models.IntegerField(choices=WEIGHT_CHOICES, default=1)

    class Meta:
        unique_together = ('metric', 'user_group')
        verbose_name = "Group Weight"

    def __str__(self):
        return f"{self.user_group} : {self.metric} ({self.factor})"