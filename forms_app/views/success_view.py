# forms_app/views/success_view.py

from django.shortcuts import render, HttpResponse
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from forms_app.models import UserReport
import os
from django.conf import settings


@login_required
def success_page(request):
    try:
        report = UserReport.objects.filter(user=request.user).latest("last_updated")
    except UserReport.DoesNotExist:
        report = None

    return render(request, "forms_app/success.html", {"report": report})


@login_required
def download_form4_file(request):
    """
    Скачивание файла формы 4: Separated_Art_Rep.xlsx
    Гарантирует принудительное скачивание
    """
    try:
        report = UserReport.objects.get(user=request.user, report_type="form4")
        file_path = os.path.join(settings.MEDIA_ROOT, report.file_path)

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Файл {file_path} не найден")

        with open(file_path, "rb") as fh:
            response = HttpResponse(
                fh.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                'attachment; filename="Separated_Art_Rep.xlsx"'
            )
            return response

    except (UserReport.DoesNotExist, FileNotFoundError) as e:
        raise PermissionDenied(
            "❌ Накопительный файл не найден. Загрузите его через форму."
        )


@login_required
def download_current_file(request):
    """
    Скачивание файла формы 5: output_stock.xlsx
    """
    try:
        report = UserReport.objects.get(user=request.user, report_type="form5")
        file_path = os.path.join(settings.MEDIA_ROOT, report.file_path)

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Файл {file_path} не найден")

        with open(file_path, "rb") as fh:
            response = HttpResponse(
                fh.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = 'attachment; filename="output_stock.xlsx"'
            return response

    except (UserReport.DoesNotExist, FileNotFoundError) as e:
        raise PermissionDenied("❌ Файл остатков не найден.")
