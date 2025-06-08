# forms_app/urls.py

from django.urls import path

# Импортируем представления
from .views.form1_view import form1
from .views.form2_view import form2
from .views.form3_view import form3
from .views.form4_view import upload_file
from forms_app.views.success_view import success_page, download_output_file

app_name = "forms_app"

urlpatterns = [
    path("form1/", form1, name="form1"),  # ← Так правильно
    path("form2/", form2, name="form2"),  #   Вызываем функции напрямую
    path("form3/", form3, name="form3"),  #   без views.
    path("upload/", upload_file, name="upload_file"),  # ✅ Новый маршрут
    path(
        "download-current/", download_output_file, name="download_output_file"
    ),  # ← новый маршрут
]
