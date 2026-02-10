from django.urls import path            # type:ignore
from . import views

urlpatterns = [
    # 1. Dashboard (The Homepage)
    path('', views.dashboard_view, name='dashboard'),

    # 2. Upload Page (To process new Excel files)
    path('upload/', views.upload_view, name='upload'),

    # 3. Live Report (The HTML table in a new tab)
    path('report/', views.report_view, name='report'),

    # 4. Summary Export (The Excel download logic)
    path('export/', views.export_view, name='export_data'),

    # 5. Detailed Export (The Detailed Excel download logic)
    path('export-detailed/', views.export_detailed_view, name='export_detailed'),

    # 6. Detailed Report (The Detailed HTML table in a new tab)
    path('report-detailed/', views.report_detailed_view, name='report_detailed'),

    path('project/<int:pk>/', views.project_detail, name='project_detail'),
]