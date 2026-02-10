from django.db import models            # type:ignore

class Project(models.Model):
    # --- ID & Metadata ---
    project_code = models.CharField(max_length=50, unique=True)
    project_name = models.CharField(max_length=255, null=True, blank=True)
    sbu = models.CharField(max_length=50, null=True, blank=True)
    stage = models.CharField(max_length=50, null=True, blank=True)
    
    # --- Dates ---
    login_date = models.DateField(null=True, blank=True)
    start_date = models.DateField(null=True, blank=True)
    end_date = models.DateField(null=True, blank=True)
    
    # --- Team ---
    sales_head = models.CharField(max_length=100, null=True, blank=True)
    sales_lead = models.CharField(max_length=100, null=True, blank=True)

    design_dh = models.CharField(max_length=100, null=True, blank=True)
    design_dm = models.CharField(max_length=100, null=True, blank=True)
    design_id = models.CharField(max_length=100, null=True, blank=True)
    design_3d = models.CharField(max_length=100, null=True, blank=True)

    ops_head = models.CharField(max_length=100, null=True, blank=True)
    ops_pm = models.CharField(max_length=100, null=True, blank=True)
    ops_om = models.CharField(max_length=100, null=True, blank=True)
    ops_ss = models.CharField(max_length=100, null=True, blank=True)
    ops_mep = models.CharField(max_length=100, null=True, blank=True)
    ops_csc = models.CharField(max_length=100, null=True, blank=True)

    # =========================================================
    # METRICS (Must match constants.py EXCEL_COL_MAP keys exactly)
    # =========================================================

    # --- Sales Metrics ---
    req_uploaded = models.FloatField(default=0.0)
    site_visit_report = models.FloatField(default=0.0)
    client_access = models.FloatField(default=0.0)
    boq_uploaded = models.FloatField(default=0.0)
    contract_uploaded = models.FloatField(default=0.0)
    boq = models.FloatField(default=0.0)         
    contract = models.FloatField(default=0.0)    

    # --- Design Metrics ---
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

    # --- Ops Metrics ---
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
    
    # Calculated Ops Metrics
    wpr_half_week = models.FloatField(default=0.0)
    manpower_ratio = models.FloatField(default=0.0)
    dpr_ratio = models.FloatField(default=0.0)
    manpower_day_ratio = models.FloatField(default=0.0)

def __str__(self):
        return self.project_code


# --- NEW: Admin Panel Structures ---

class Department(models.Model):
    name = models.CharField(max_length=50, unique=True) # e.g., 'Sales', 'Design', 'Operations'

    def __str__(self):
        return self.name

class UserGroup(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    name = models.CharField(max_length=50) # e.g., 'ID', 'DH', 'PM', 'Sales Lead'

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

    name = models.CharField(max_length=50, unique=True) # e.g. "Completeness"
    color = models.CharField(max_length=20, choices=COLOR_CHOICES, default='secondary') # e.g. "success", "warning", "danger"

    def __str__(self):
        return self.name

class Metric(models.Model):
    STAGE_CHOICES = [('Pre', 'Pre-Stage'), ('Post', 'Post-Stage')]

    label = models.CharField(max_length=100) # The display name e.g., "Key Plans Ratio"
    field_name = models.CharField(max_length=100, help_text="Must match a field in the Project model exactly (e.g., key_plans_ratio)")
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    stage = models.CharField(max_length=10, choices=STAGE_CHOICES)
    
    # Logic & Config
    default_threshold = models.FloatField(default=0.0)

    # Success Category (For color-coding in UI)
    success_metric = models.ForeignKey(SuccessMetric, on_delete=models.SET_NULL, null=True, blank=True)
    
    # User Group Link (Which roles see this metric?)
    # ManyToMany because "Key Plans" might be visible to both 'ID' and 'DH'
    visible_to_groups = models.ManyToManyField(UserGroup, blank=True)

    # Credit System (Phase 2 Prep)
    is_manual_credit = models.BooleanField(default=False)
    credit_weight = models.FloatField(default=0.0) # Auto-calculated

    def __str__(self):
        return f"{self.label} ({self.department})"