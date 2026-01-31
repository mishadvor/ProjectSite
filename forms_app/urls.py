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
    clear_form4_by_date,
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

# Форма 8
from .views.form8_view import (
    form8_upload,
    form8_clear,
    form8_export,
    form8_clear_by_date,
)

# Форма 9
from .views.form9_view import form9_view

# Форма 10
from .views.form10_view import form10_view

# --- Форма 11 ---
from .views.form11_view import form11_view

# ---- Форма 12 -----
from .views.form12_view import (
    upload_file12,
    form12_list,
    form12_detail,
    form12_edit,
    form12_delete,
    form12_delete_article,
    form12_delete_all,
    form12_delete_by_date,
    export_form12_excel,
    form12_chart,
    clear_form12_data,
)
from .views.form13_view import form13_simple_upload
from .views.form14_view import (
    upload_file14,
    form14_list,
    form14_chart,
    clear_form14_data,
    form14_delete_by_date,
    export_form14_excel,
    form14_api_data,
)
from .views.form15_view import (
    form15_view,
    form15_edit_pattern,
    form15_delete_pattern,
    form15_calculate,
    form15_clear_all,
    form15_import_excel,
)
from .views.form16_view import (
    form16_main,  # Это главная страница
    form16_edit_table,
    form16_generate_report,
    form16_delete_all,
)
from .views.form17_view import (
    form17_view,
    form17_load_chart,
    form17_delete_chart,
)

# --- Форма 18: Себестоимость артикулов ---
from .views.form18_view import (
    form18_list,
    form18_edit,
    form18_delete,
)

from .views.form19_view import (
    form19_view,
)


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
    path("form4/clear-by-date/", clear_form4_by_date, name="form4_clear_by_date"),
    # Графики — только один раз!
    path("form4/<str:code>/chart/", form4_chart, name="form4_chart"),
    path(
        "form4/<str:code>/chart/<str:chart_type>/", form4_chart, name="form4_chart_type"
    ),
    # Редактирование
    path("form4/edit/<int:pk>/", form4_edit, name="form4_edit"),
    # Детали (в самом конце!)
    path("form4/<str:code>/", form4_detail, name="form4_detail"),
    # --- Форма 8 ---
    path("form8/", form8_upload, name="form8_upload"),
    path("form8/clear/", form8_clear, name="form8_clear"),
    path("form8/export/", form8_export, name="form8_export"),
    path("form8/clear-by-date/", form8_clear_by_date, name="form8_clear_by_date"),
    # -----Форма 9 ------
    path("form9/", form9_view, name="form9_view"),
    # -----Форма 10 ------
    path("form10/", form10_view, name="form10_view"),
    # --- Форма 11 -----
    path("form11/", form11_view, name="form11_view"),
    # --- Форма 12 ---
    path("form12/upload/", upload_file12, name="form12_upload"),
    path("form12/list/", form12_list, name="form12_list"),
    path("form12/detail/<str:wb_article>/", form12_detail, name="form12_detail"),
    path("form12/edit/<int:pk>/", form12_edit, name="form12_edit"),
    path("form12/export/", export_form12_excel, name="form12_export"),
    path(
        "form12/chart/<str:wb_article>/<str:chart_type>/",
        form12_chart,
        name="form12_chart",
    ),
    path("form12/clear/", clear_form12_data, name="form12_clear"),
    # Три уровня удаления:
    path("form12/delete/<int:pk>/", form12_delete, name="form12_delete"),  # Одна запись
    path(
        "form12/delete-article/<str:wb_article>/",
        form12_delete_article,
        name="form12_delete_article",
    ),  # Один артикул
    path(
        "form12/delete-all/", form12_delete_all, name="form12_delete_all"
    ),  # Все данные
    path("form12/delete-by-date/", form12_delete_by_date, name="form12_delete_by_date"),
    # --- Форма 13 (простая версия) ---
    path("form13/", form13_simple_upload, name="form13_simple"),
    # --- Форма 14 (Агрегированные данные по всем артикулам) ---
    path("form14/upload/", upload_file14, name="form14_upload"),
    path("form14/", form14_list, name="form14_list"),
    path("form14/chart/<str:chart_type>/", form14_chart, name="form14_chart"),
    path("form14/chart/", form14_chart, name="form14_chart_default"),
    path("form14/clear/", clear_form14_data, name="form14_clear"),
    path("form14/delete-by-date/", form14_delete_by_date, name="form14_delete_by_date"),
    path("form14/export/", export_form14_excel, name="form14_export"),
    path("form14/api/<str:chart_type>/", form14_api_data, name="form14_api_data"),
    path("form15/", form15_view, name="form15_view"),
    path("form15/edit/<int:pk>/", form15_edit_pattern, name="form15_edit_pattern"),
    path(
        "form15/delete/<int:pk>/", form15_delete_pattern, name="form15_delete_pattern"
    ),
    path("form15/calculate/", form15_calculate, name="form15_calculate"),
    path("form15/clear-all/", form15_clear_all, name="form15_clear_all"),
    path("form15/import-excel/", form15_import_excel, name="form15_import_excel"),
    # --- Форма 16 ---
    path("form16/", form16_main, name="form16_main"),
    path("form16/edit/", form16_edit_table, name="form16_edit_table"),
    path("form16/generate/", form16_generate_report, name="form16_generate_report"),
    path("form16/delete-all/", form16_delete_all, name="form16_delete_all"),
    # --- Форма 17 ----
    path("form17/", form17_view, name="form17_view"),
    path("form17/load/<int:pk>/", form17_load_chart, name="form17_load"),
    path("form17/delete/<int:pk>/", form17_delete_chart, name="form17_delete"),
    # --- Форма 18 ----
    path("form18/", form18_list, name="form18_list"),
    path("form18/edit/<int:pk>/", form18_edit, name="form18_edit"),
    path("form18/delete/<int:pk>/", form18_delete, name="form18_delete"),
    # --- Форма 19 ---
    path("form19/", form19_view, name="form19_view"),
]
