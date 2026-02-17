from django.contrib import admin            # type: ignore
from .models import Project, Metric, Department, UserGroup, SuccessMetric, MetricWeight

# --- 1. Success Metrics ---
@admin.register(SuccessMetric)
class SuccessMetricAdmin(admin.ModelAdmin):
    list_display = ('name', 'color')
    list_editable = ('color',)

# --- 2. Organization Structure ---
@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name',)

@admin.register(UserGroup)
class UserGroupAdmin(admin.ModelAdmin):
    list_display = ('name', 'department')
    list_filter = ('department',)
    search_fields = ('name',)

# --- 3. Weight Configuration Inline ---
class MetricWeightInline(admin.TabularInline):
    model = MetricWeight
    extra = 0
    min_num = 0
    can_delete = True
    verbose_name = "Weight per Group"
    verbose_name_plural = "Scoring Weights (Higher = More Points)"
    autocomplete_fields = ['user_group']

# --- 4. Project Data ---
@admin.register(Project)
class ProjectAdmin(admin.ModelAdmin):
    list_display = ('project_code', 'project_name', 'sbu', 'stage', 'login_date')
    list_filter = ('sbu', 'stage', 'login_date')
    search_fields = ('project_code', 'project_name', 'sales_lead', 'ops_pm')
    date_hierarchy = 'login_date'

# --- 5. Metrics Configuration ---
@admin.register(Metric)
class MetricAdmin(admin.ModelAdmin):
    list_display = ('label', 'stage', 'department', 'min_threshold', 'max_threshold')
    list_editable = ('min_threshold', 'max_threshold')
    list_filter = ('department', 'stage', 'success_metric')
    search_fields = ('label', 'field_name')
    
    # This places the Weight Table directly inside the Metric page
    inlines = [MetricWeightInline]

    # Organize the detail view
    fieldsets = (
        ('Basic Info', {
            'fields': ('label', 'field_name', 'department', 'stage')
        }),
        ('Scoring Logic', {
            'fields': ('min_threshold', 'max_threshold'),
            'description': '<br><b>Min:</b> Hurdle to start earning points. <br><b>Max:</b> Cap for maximum points.'
        }),
        ('Visibility', {
            'fields': ('visible_to_groups', 'success_metric')
        }),
    )

    filter_horizontal = ('visible_to_groups',)

    @admin.display(description="Active Weights")
    def get_assigned_weights(self, obj):
        # formatted summary for list view
        weights = obj.metricweight_set.filter(factor__gt=0)
        if not weights.exists():
            return "-"
        return ", ".join([f"{w.user_group.name} ({w.factor})" for w in weights])