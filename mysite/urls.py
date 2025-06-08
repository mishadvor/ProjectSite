# myproject/urls.py

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from forms_app.views.form4_view import upload_file  # ✅ Так должно быть
from forms_app.views.success_view import success_page


urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("accounts.urls")),
    path("", include("home.urls")),  # ← путь к приложению - Домашняя страница
    path("", include("forms_app.urls")),  # путь к формам
    path(
        "theory/", include("theory.urls", namespace="theory")
    ),  # путь к разделу-приложению - Теория
    path("ckeditor/", include("ckeditor_uploader.urls")),  # ← маршруты для CKEditor
    path("upload/", upload_file, name="upload_file"),
    path("success/", success_page, name="success_page"),
]

# Не забудь подключение медиафайлов во время разработки
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
