# forms_app/views/dashboard_view.py
from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.db.utils import OperationalError
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import io, base64
import pandas as pd
from datetime import timedelta


@login_required
def dashboard(request):
    user = request.user
    today = timezone.now().date()

    # === 📅 ОПРЕДЕЛЯЕМ ЦЕЛЕВУЮ ДАТУ ===
    target_date = today
    date_label = "на сегодня"

    try:
        from forms_app.models import Form14Data

        # Если за сегодня данных нет → берём последнюю доступную дату
        if not Form14Data.objects.filter(user=user, date=today).exists():
            latest = Form14Data.objects.filter(user=user).order_by("-date").first()
            if latest:
                target_date = latest.date
                diff = (today - target_date).days
                if diff == 1:
                    date_label = f"на {target_date.strftime('%d.%m.%Y')} (вчера)"
                else:
                    date_label = f"на {target_date.strftime('%d.%m.%Y')}"
    except (ImportError, OperationalError):
        pass

    # === 📊 МЕТРИКИ ИЗ FORM14 ===
    metrics = {
        "orders": 0,  # Заказы
        "order_amount": 0,  # Сумма заказов
        "sold": 0,  # Выкуплено
        "transfer": 0,  # К перечислению
        "stock": 0,  # Остаток
    }
    charts = {}

    try:
        from forms_app.models import Form14Data

        record = Form14Data.objects.filter(user=user, date=target_date).first()
        if record:
            metrics["orders"] = int(record.total_orders_qty or 0)
            metrics["order_amount"] = round(
                float(record.total_order_amount_net or 0), 1
            )
            metrics["sold"] = int(record.total_sold_qty or 0)
            metrics["transfer"] = round(float(record.total_transfer_amount or 0), 1)
            metrics["stock"] = int(record.total_current_stock or 0)
    except (ImportError, OperationalError, AttributeError):
        pass

    # === 📈 ГРАФИК: Заказы vs Выкуплено (7 дней) ===
    try:
        from forms_app.models import Form14Data

        week_ago = target_date - timedelta(days=6)
        weekly_qs = Form14Data.objects.filter(
            user=user, date__range=[week_ago, target_date]
        ).order_by("date")

        if weekly_qs.exists():
            # Готовим данные
            data = [
                {
                    "date": r.date,
                    "orders": r.total_orders_qty or 0,
                    "sold": r.total_sold_qty or 0,
                }
                for r in weekly_qs
            ]
            df = pd.DataFrame(data)
            df["date_str"] = pd.to_datetime(df["date"]).dt.strftime("%d.%m")

            buf = io.BytesIO()
            plt.figure(figsize=(8, 4))

            # Линия 1: Заказы (синяя)
            plt.plot(
                df["date_str"],
                df["orders"],
                marker="o",
                color="#007bff",
                linewidth=2,
                label="Заказы",
                markersize=5,
            )
            # Линия 2: Выкуплено (зелёная)
            plt.plot(
                df["date_str"],
                df["sold"],
                marker="s",
                color="#28a745",
                linewidth=2,
                label="Выкуплено",
                markersize=5,
            )

            plt.title(
                "📦 Заказы vs 🛍️ Выкуплено (7 дней)", fontsize=12, fontweight="bold"
            )
            plt.xlabel("Дата")
            plt.ylabel("Количество, шт.")
            plt.xticks(rotation=45, ha="right")
            plt.legend(fontsize=9)
            plt.grid(alpha=0.3, linestyle="--")
            plt.tight_layout()
            plt.savefig(buf, format="png", dpi=100, bbox_inches="tight")
            buf.seek(0)
            charts["trend_7d"] = base64.b64encode(buf.read()).decode("utf-8")
            plt.close()
    except Exception:
        pass

    context = {
        "page_title": "📊 Дашборд",
        "date_label": date_label,
        "today": target_date,
        "metrics": metrics,
        "charts": charts,
    }
    return render(request, "forms_app/dashboard.html", context)
