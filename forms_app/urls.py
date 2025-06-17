# forms_app/urls.py

from django.urls import path

# Импортируем представления
from .views.form1_view import form1
from .views.form2_view import form2
from .views.form3_view import form3
from .views.form4_view import upload_file
from .views.form5_view import form5
from .views.success_view import success_page, download_form4_file
from .views.success_view import download_current_file
from .views.reports_view import my_reports

app_name = "forms_app"

urlpatterns = [
    path("form1/", form1, name="form1"),
    path("form2/", form2, name="form2"),
    path("form3/", form3, name="form3"),
    path("upload/", upload_file, name="upload_file"),
    path("form5/", form5, name="form5"),
    # Страница успеха
    path("success/", success_page, name="success_page"),
    # Скачивание файлов
    path(
        "download-output/", download_form4_file, name="download_output_file"
    ),  # форма 4
    path(
        "download-current/", download_current_file, name="download_current_file"
    ),  # форма 5
    # Мои отчёты
    path("my-reports/", my_reports, name="my_reports"),
]
