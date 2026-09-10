"""Settings prove their own defaults: SQLite unless told otherwise; media outside the repository."""
from cw import settings


def test_database_defaults_to_sqlite():
    config = settings.database_from_env({})
    assert config["ENGINE"] == "django.db.backends.sqlite3"
    assert str(config["NAME"]).endswith("dev.sqlite3")


def test_database_url_selects_postgresql():
    config = settings.database_from_env({"DATABASE_URL": "postgres://u:p@localhost:5432/cw"})
    assert config["ENGINE"] == "django.db.backends.postgresql"
    assert config["NAME"] == "cw"


def test_media_root_outside_repo():
    root = settings.media_root_from_env({})
    assert not root.is_relative_to(settings.BASE_DIR)
    custom = settings.media_root_from_env({"MEDIA_ROOT": "/srv/cw_media"})
    assert str(custom) == "/srv/cw_media"
