# forms_app/views/reports_view.py

from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from forms_app.models import UserReport
import os
from datetime import datetime
from django.utils import timezone
from django.conf import settings


@login_required
def my_reports(request):
    user = request.user

    # === Форма 4: накопительные отчёты ===
    form4_reports = UserReport.objects.filter(user=user, report_type="form4").order_by(
        "-last_updated"
    )

    # === Форма 5: остатки ===
    form5_report = UserReport.objects.filter(user=user, report_type="form5").first()

    # === Проверка наличия файла формы 5 ===
    stock_path = os.path.join(
        settings.MEDIA_ROOT, "user_stock", str(user.id), "output_stock.xlsx"
    )
    stock_exists = os.path.exists(stock_path)

    last_updated = None
    if stock_exists:
        try:
            last_updated = datetime.fromtimestamp(os.path.getmtime(stock_path))
        except Exception as e:
            print(f"❌ Ошибка при получении даты файла: {e}")

    return render(
        request,
        "forms_app/my_reports.html",
        {
            "form4_reports": form4_reports,
            "form5_report": form5_report,
            "stock_exists": stock_exists,
            "stock_last_updated": last_updated,
            "user_id": user.id,
        },
    )
