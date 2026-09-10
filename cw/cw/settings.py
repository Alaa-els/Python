"""Settings for the Commercial Workbench.

Every value that varies by machine or environment comes from the environment
(a .env file is read if present). Nothing product-specific lives here.
"""
import os
import sys
from pathlib import Path

import dj_database_url
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

RUNNING_TESTS = "pytest" in sys.modules

DEBUG = os.environ.get("DEBUG", "1") == "1"

SECRET_KEY = os.environ.get("SECRET_KEY", "")
if not SECRET_KEY:
    if DEBUG or RUNNING_TESTS:
        SECRET_KEY = "development-only-key-not-for-use-outside-a-developer-machine"
    else:
        raise RuntimeError("SECRET_KEY must be set in the environment when DEBUG is off")

ALLOWED_HOSTS = [h for h in os.environ.get("ALLOWED_HOSTS", "localhost,127.0.0.1").split(",") if h]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "simple_history",
    "django_htmx",
    "core",
    "projects",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
    "django_htmx.middleware.HtmxMiddleware",
    "simple_history.middleware.HistoryRequestMiddleware",
]

ROOT_URLCONF = "cw.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "cw.wsgi.application"


def database_from_env(environ):
    """Return the DATABASES["default"] mapping for the given environment.

    DATABASE_URL selects the engine; SQLite in the repository folder is the
    default. Under pytest, TEST_DATABASE_URL wins when it is set, so the same
    suite can run against PostgreSQL without touching the developer database.
    """
    default_url = "sqlite:///" + str(BASE_DIR / "dev.sqlite3")
    url = environ.get("DATABASE_URL") or default_url
    if RUNNING_TESTS and environ.get("TEST_DATABASE_URL"):
        url = environ["TEST_DATABASE_URL"]
    return dj_database_url.parse(url, conn_max_age=0)


DATABASES = {"default": database_from_env(os.environ)}

AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]

LANGUAGE_CODE = "en-gb"
TIME_ZONE = "Europe/London"
USE_I18N = True
USE_TZ = True

STATIC_URL = "static/"


def media_root_from_env(environ):
    """Uploaded files live outside the repository unless MEDIA_ROOT says otherwise."""
    return Path(environ.get("MEDIA_ROOT") or (BASE_DIR.parent / "cw_media"))


MEDIA_ROOT = media_root_from_env(os.environ)
MEDIA_URL = "media/"

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"
