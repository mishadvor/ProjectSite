# forms_app/urls.py

from django.urls import path

# Импортируем представления
from .views.form1_view import form1
from .views.form2_view import form2
from .views.form3_view import form3
from .views.form4_view import (
    upload_file,
    form4_list,
    form4_detail,
    form4_edit,
    form4_chart,
    export_form4_excel,
    clear_form4_data,
)

from .views.form5_view import form5

# Страница успеха и загрузка файлов
from .views.success_view import success_page, download_form4_file, download_current_file

# Предпросмотр и замена остатков
from .views.stock_replace_view import replace_stock, preview_output_stock

# Мои отчёты
from .views.reports_view import my_reports

# Форма 6
from forms_app.views.form6_view import form6
from forms_app.views.form6_sql_views import preview_sql, download_sql
from .views.form6_sql_views import (
    editable_preview_sql,
    save_stock_sql,
)
from forms_app.views.stock_replace_view import replace_sql_stock
from forms_app.views.form6_sql_views import reset_stock_sql

# Форма 7
from .views.form7_view import form7_upload, form7_graph, clear_form7_data

app_name = "forms_app"

urlpatterns = [
    # --- Основные формы ---
    path("form1/", form1, name="form1"),
    path("form2/", form2, name="form2"),
    path("form3/", form3, name="form3"),
    path("form5/", form5, name="form5"),
    # --- Страница успеха ---
    path("success/", success_page, name="success_page"),
    # --- Скачивание файлов ---
    path("download-current/", download_current_file, name="download_current_file"),
    # --- Мои отчёты ---
    path("my-reports/", my_reports, name="my_reports"),
    # --- Замена и предпросмотр остатков (форма 5) ---
    path("form5/replace_stock/", replace_stock, name="replace_stock"),
    path("form5/preview/", preview_output_stock, name="preview_output_stock"),
    # --- Форма 6 ---
    path("form6/", form6, name="form6"),
    path("form6/preview/", preview_sql, name="preview_sql"),
    path("form6/download/", download_sql, name="download_sql"),
    path("form6/edit/", editable_preview_sql, name="editable_preview_sql"),
    path("form6/save/", save_stock_sql, name="save_stock_sql"),
    path("form6/replace_sql/", replace_sql_stock, name="replace_sql_stock"),
    path("form6/reset/", reset_stock_sql, name="reset_stock_sql"),
    # --- Форма 7 ---
    path("form7/upload/", form7_upload, name="form7_upload"),
    path("form7/graph/", form7_graph, name="form7_graph"),
    path("form7/clear/", clear_form7_data, name="clear_form7_data"),
    # --- Form4 (SQL) — Сначала конкретные, потом общие ---
    path("form4/upload/", upload_file, name="form4_upload"),
    path("form4/", form4_list, name="form4_list"),
    # Экспорт и очистка
    path("form4/export/", export_form4_excel, name="form4_export"),
    path("form4/clear/", clear_form4_data, name="form4_clear"),
    # Графики
    path(
        "form4/<str:code>/chart/<str:chart_type>/", form4_chart, name="form4_chart_type"
    ),
    path("form4/<str:code>/chart/", form4_chart, name="form4_chart"),
    # Редактирование
    path("form4/edit/<int:pk>/", form4_edit, name="form4_edit"),
    # Детали (в самом конце!)
    path("form4/<str:code>/", form4_detail, name="form4_detail"),
]
