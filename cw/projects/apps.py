from django.apps import AppConfig


class ProjectsConfig(AppConfig):
    """Scheme register and everything that hangs off a scheme (models arrive at S02 and S04)."""

    default_auto_field = "django.db.models.BigAutoField"
    name = "projects"
