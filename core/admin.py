from django.contrib import admin            # type: ignore
from django.db.models import Sum            # type: ignore
from .models import Project, Metric, Department, UserGroup, SuccessMetric, MetricWeight

# --- 1. Success Metrics (Tags) ---
@admin.register(SuccessMetric)
class SuccessMetricAdmin(admin.ModelAdmin):
    list_display = ('name', 'color')
    list_editable = ('color',)

# --- 2. User Groups & Departments ---
@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name',)

@admin.register(UserGroup)
class UserGroupAdmin(admin.ModelAdmin):
    list_display = ('name', 'department')
    list_filter = ('department',)
    search_fields = ('name',)

# --- 3. Inline for Metric Weights ---
class MetricWeightInline(admin.TabularInline):
    model = MetricWeight
    extra = 0 # Clean UI: Start with no empty rows
    min_num = 0
    can_delete = True
    verbose_name = "Weight per Group"
    verbose_name_plural = "Weight Configuration (Auto-Balances 100%)"
    
    # This creates a searchable dropdown for Groups (Best for long lists)
    autocomplete_fields = ['user_group']

# --- 4. Projects (Standard) ---
@admin.register(Project)
class ProjectAdmin(admin.ModelAdmin):
    list_display = ('project_name', 'project_code', 'stage', 'sbu', 'login_date')
    search_fields = ('project_name', 'project_code')
    # REMOVED 'department' from here to fix the error
    list_filter = ('sbu', 'stage')

# --- 5. The Smart Metric Admin (Auto-Balancing) ---
@admin.register(Metric)
class MetricAdmin(admin.ModelAdmin):
    list_display = ('label', 'department', 'stage', 'get_assigned_weights', 'success_metric')
    list_filter = ('department', 'stage')
    search_fields = ('label', 'field_name')
    
    # THIS puts the Weight Table inside the Metric Page
    inlines = [MetricWeightInline]

    fieldsets = (
        ('Basic Info', {
            'fields': ('label', 'field_name', 'department', 'stage', 'success_metric')
        }),
        ('Logic', {
            'fields': ('default_threshold',) 
        }),
        ('Visibility Filter', {
            'fields': ('visible_to_groups',),
            'description': 'Use this to tag groups broadly. Use the table below to set specific weights.'
        }),
    )

    filter_horizontal = ('visible_to_groups',)

    @admin.display(description="Configured Weights")
    def get_assigned_weights(self, obj):
        # Shows a summary in the list view: "ID (10), DM (5)"
        return ", ".join([f"{w.user_group.name}: {w.factor}" for w in obj.metricweight_set.all()])