"""The configured database accepts a row and returns it (a round trip, so the engine is really exercised)."""
import pytest
from django.contrib.auth import get_user_model
from django.db import connection


@pytest.mark.django_db
def test_database_accepts_a_row():
    user_model = get_user_model()
    user_model.objects.create_user(username="sample-user", password="sample-password-not-real")
    assert user_model.objects.filter(username="sample-user").count() == 1


def test_report_engine_in_use():
    assert connection.vendor in {"sqlite", "postgresql"}
