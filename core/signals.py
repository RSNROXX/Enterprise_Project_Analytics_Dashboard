# core/signals.py
from django.dispatch import receiver                                # type: ignore
from django.db.models.signals import post_save, m2m_changed         # type: ignore
from .models import Metric

# OLD LOGIC REMOVED.
# The new system uses the MetricWeight table and calculates percentages live on the dashboard. No signals are required here.