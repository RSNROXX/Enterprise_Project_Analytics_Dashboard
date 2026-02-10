from django.apps import AppConfig                       # type: ignore

class CoreConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'core'  # Make sure this matches your folder name (it might be 'flipspaces_analytics' or 'core')

    def ready(self):
        import core.signals  # <--- THIS IS THE KEY LINE