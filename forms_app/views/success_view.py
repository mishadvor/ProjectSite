# forms_app/views/success_view.py

from django.shortcuts import render
from django.http import HttpResponse
from forms_app.models import UserReport
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
import os
from django.conf import settings


def success_page(request):
    try:
        report = UserReport.objects.filter(user=request.user).latest("last_updated")
    except UserReport.DoesNotExist:
        report = None

    return render(request, "forms_app/success.html", {"report": report})


# success_view.py

from django.http import HttpResponse
from forms_app.models import UserReport  # ✅ Так будет работать
from django.core.exceptions import PermissionDenied
from django.contrib.auth.decorators import login_required
import os


@login_required
def download_output_file(request):
    try:
        report = UserReport.objects.get(user=request.user)
        output_file_path = report.output_file.path
    except (UserReport.DoesNotExist, FileNotFoundError):
        raise PermissionDenied("❌ Файл не найден. Загрузите отчет.")

    with open(output_file_path, "rb") as fh:
        response = HttpResponse(fh.read(), content_type="application/octet-stream")
        response["Content-Disposition"] = (
            'attachment; filename="Separated_Art_Rep.xlsx"'
        )
        return response
