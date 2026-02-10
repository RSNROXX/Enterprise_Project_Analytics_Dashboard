from django.contrib import admin            # type: ignore
from django.db.models import Sum            # type: ignore
from .models import Project, Metric, Department, UserGroup, SuccessMetric

# --- 1. Success Metrics (Tags) ---
@admin.register(SuccessMetric)
class SuccessMetricAdmin(admin.ModelAdmin):
    list_display = ('name', 'color')
    list_editable = ('color',)

# --- 2. User Groups & Departments ---
admin.site.register(UserGroup)
admin.site.register(Department)

# --- 3. Projects (Standard) ---
@admin.register(Project)
class ProjectAdmin(admin.ModelAdmin):
    list_display = ('project_name', 'project_code', 'stage', 'sbu', 'login_date')
    search_fields = ('project_name', 'project_code')
    # REMOVED 'department' from here to fix the error
    list_filter = ('sbu', 'stage')

# --- 4. The Smart Metric Admin (Auto-Balancing) ---
@admin.register(Metric)
class MetricAdmin(admin.ModelAdmin):
    list_display = ('label', 'department', 'stage', 'get_groups', 'credit_weight', 'is_manual_credit', 'success_metric')
    list_filter = ('department', 'stage', 'visible_to_groups')
    list_editable = ('credit_weight', 'is_manual_credit')
    filter_horizontal = ('visible_to_groups',)
    
    fieldsets = (
        ('Basic Info', {
            'fields': ('label', 'field_name', 'department', 'stage', 'visible_to_groups')
        }),
        ('Logic & Thresholds', {
            'fields': ('default_threshold', 'success_metric') 
        }),
        ('Credit System', {
            'fields': ('is_manual_credit', 'credit_weight'),
            'description': 'Check "Manual" to lock this number. Uncheck to let the system auto-balance it to reach 100%.'
        }),
    )

    @admin.display(description='User Groups')
    def get_groups(self, obj):
        return ", ".join([g.name for g in obj.visible_to_groups.all()])

    def save_related(self, request, form, formsets, change):
        """
        Triggered AFTER the Many-to-Many 'visible_to_groups' are saved.
        This is the safe place to run calculations.
        """
        super().save_related(request, form, formsets, change)
        
        # Run the re-balancing logic
        obj = form.instance
        self.rebalance_credits(obj)

    def rebalance_credits(self, current_metric):
        """
        The Logic:
        For every group this metric touches, ensure that Dept+Stage+Group sums to 100.
        """
        department = current_metric.department
        stage = current_metric.stage
        groups = current_metric.visible_to_groups.all()

        if not groups:
            return 

        # Iterate over every group to ensure that "ID Pre-Sales" sums to 100 
        # AND "DM Pre-Sales" sums to 100, even if they share this metric.
        for group in groups:
            # 1. Get the "Team" (All metrics for this specific bucket)
            team_metrics = Metric.objects.filter(
                department=department, 
                stage=stage, 
                visible_to_groups=group
            ).distinct()

            # 2. Separate into "Fixed" (Manual) and "Flexible" (Auto)
            manuals = team_metrics.filter(is_manual_credit=True)
            autos = team_metrics.filter(is_manual_credit=False)

            # 3. Calculate the available pool
            # (Handle None result if no manuals exist)
            manual_sum = manuals.aggregate(Sum('credit_weight'))['credit_weight__sum'] or 0
            remaining_pool = 100 - manual_sum
            
            # 4. Distribute the pool
            auto_count = autos.count()
            
            if auto_count > 0:
                # Calculate new weight (e.g., 80 / 4 = 20)
                # Use max(0, ...) to prevent negative scores if Manuals > 100
                new_weight = max(0, remaining_pool / auto_count)
                
                # 5. Apply the update
                autos.update(credit_weight=new_weight)