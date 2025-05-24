# myproject/urls.py

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static


urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("home.urls")),  # ← Эта строка важна!
    path("", include("forms_app.urls")),
    path("theory/", include("theory.urls", namespace="theory")),
    # ⚙️ Добавь это!
    path("ckeditor/", include("ckeditor_uploader.urls")),  # ← маршруты для CKEditor
]

# Не забудь подключение медиафайлов во время разработки
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
