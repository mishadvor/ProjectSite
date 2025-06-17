# myproject/urls.py

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from forms_app.views.success_view import success_page


urlpatterns = [
    path("admin/", admin.site.urls),
    path("accounts/", include("accounts.urls", namespace="accounts")),
    path(
        "forms/", include("forms_app.urls")
    ),  # Использует app_name из forms_app/urls.py
    path("theory/", include("theory.urls", namespace="theory")),
    path("", include("home.urls")),
    path("ckeditor/", include("ckeditor_uploader.urls")),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
