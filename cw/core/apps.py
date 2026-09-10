from django.apps import AppConfig


class CoreConfig(AppConfig):
    """Companies, users, roles, membership, settings and history helpers (models arrive at S02)."""

    default_auto_field = "django.db.models.BigAutoField"
    name = "core"
