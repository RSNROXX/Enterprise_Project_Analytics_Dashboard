from django.db.models.signals import post_save, m2m_changed, post_delete                # type: ignore
from django.dispatch import receiver                                                    # type: ignore
from django.db import transaction                                                       # type: ignore
from .models import Metric, UserGroup

def distribute_group_credits(group):
    """
    The Brain: Balances credits for a specific User Group to ensure Sum = 100.
    """
    if not group:
        return

    # 1. Fetch all metrics in this group
    metrics = group.metric_set.all()
    if not metrics.exists():
        return

    # 2. Separate Manual vs Auto
    # We use a list to avoid DB locking issues during updates
    manual_metrics = [m for m in metrics if m.is_manual_credit]
    auto_metrics = [m for m in metrics if not m.is_manual_credit]

    # 3. Calculate Totals
    manual_sum = sum(m.credit_weight for m in manual_metrics)
    remaining_credit = 100.0 - manual_sum
    
    # Safety: If manual exceeds 100, we cap remaining at 0
    if remaining_credit < 0:
        remaining_credit = 0

    # 4. Distribute Remaining Credits
    auto_count = len(auto_metrics)
    
    if auto_count > 0:
        new_weight = round(remaining_credit / auto_count, 2)
        
        # Bulk Update the Auto Metrics (efficiently)
        # We filter the QuerySet again to use the .update() method
        group.metric_set.filter(is_manual_credit=False).update(credit_weight=new_weight)

# --- SIGNALS ---

@receiver(m2m_changed, sender=Metric.visible_to_groups.through)
def on_group_assignment_change(sender, instance, action, reverse, model, pk_set, **kwargs):
    """
    Triggered when a Metric is added to or removed from a User Group.
    """
    if action in ["post_add", "post_remove", "post_clear"]:
        # If we are changing the groups on a Metric instance
        if isinstance(instance, Metric):
            # Recalculate for all groups this metric is now part of (or was part of)
            if pk_set:
                groups = UserGroup.objects.filter(pk__in=pk_set)
                for group in groups:
                    distribute_group_credits(group)
            # Also recalc for the groups currently assigned (in case of removal)
            for group in instance.visible_to_groups.all():
                distribute_group_credits(group)

        # If we are changing metrics on a UserGroup instance (Reverse)
        elif isinstance(instance, UserGroup):
            distribute_group_credits(instance)

@receiver(post_save, sender=Metric)
def on_metric_save(sender, instance, created, **kwargs):
    """
    Triggered when a Metric's 'is_manual_credit' or value is changed.
    """
    # We need to recalculate every group this metric belongs to
    # Use transaction.on_commit to ensure M2M relations are ready if it's a new object
    def _do_recalc():
        for group in instance.visible_to_groups.all():
            distribute_group_credits(group)
    
    transaction.on_commit(_do_recalc)