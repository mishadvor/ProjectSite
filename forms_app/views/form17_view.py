# forms_app/views/form17_view.py
import io
import base64
import math
from datetime import datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseBadRequest
from forms_app.models import ManualChart, ManualChartDataPoint


@login_required
def form17_view(request):
    charts = (
        ManualChart.objects.filter(user=request.user)
        .prefetch_related("data_points")
        .order_by("-created_at")
    )
    chart_data = None
    table_data = []
    current_title = ""
    current_label1 = "Значение 1"
    current_label2 = ""
    loaded_chart_id = None

    if request.method == "POST":
        action = request.POST.get("action", "save")
        current_title = request.POST.get("title", "").strip()
        current_label1 = (
            request.POST.get("label1", "Значение 1").strip() or "Значение 1"
        )
        current_label2 = request.POST.get("label2", "").strip()
        chart_id = request.POST.get("chart_id")
        dates_raw = request.POST.getlist("date")
        values1_raw = request.POST.getlist("value1")
        values2_raw = request.POST.getlist("value2")

        cleaned = [
            (d.strip(), v1.strip(), v2.strip() if v2 else "")
            for d, v1, v2 in zip(dates_raw, values1_raw, values2_raw)
            if d.strip() and v1.strip()
        ]

        if not current_title:
            messages.error(request, "Укажите название графика.")
            return render(
                request,
                "forms_app/form17.html",
                {
                    "charts": charts,
                    "table_data": [],
                    "current_title": current_title,
                    "current_label1": current_label1,
                    "current_label2": current_label2,
                    "loaded_chart_id": loaded_chart_id,
                },
            )

        if not cleaned:
            messages.error(request, "Нет данных для сохранения или отображения.")
            return render(
                request,
                "forms_app/form17.html",
                {
                    "charts": charts,
                    "table_build": [],
                    "current_title": current_title,
                    "current_label1": current_label1,
                    "current_label2": current_label2,
                    "loaded_chart_id": loaded_chart_id,
                },
            )

        try:
            parsed_data = []
            for d_str, v1_str, v2_str in cleaned:
                date_obj = datetime.strptime(d_str, "%Y-%m-%d").date()
                val1 = float(v1_str)
                val2 = float(v2_str) if v2_str else None
                parsed_data.append((date_obj, val1, val2))
            parsed_data.sort(key=lambda x: x[0])
        except ValueError as e:
            messages.error(request, f"Ошибка в данных: {e}")
            return render(
                request,
                "forms_app/form17.html",
                {
                    "charts": charts,
                    "table_data": [],
                    "current_title": current_title,
                    "current_label1": current_label1,
                    "current_label2": current_label2,
                    "loaded_chart_id": loaded_chart_id,
                },
            )

        if action == "save":
            if chart_id:
                chart = get_object_or_404(ManualChart, pk=chart_id, user=request.user)
                chart.title = current_title
                chart.label1 = current_label1
                chart.label2 = current_label2
                chart.save()
                chart.data_points.all().delete()
            else:
                chart = ManualChart.objects.create(
                    title=current_title,
                    label1=current_label1,
                    label2=current_label2,
                    user=request.user,
                )

            for date_val, val1, val2 in parsed_data:
                ManualChartDataPoint.objects.create(
                    chart=chart,
                    date=date_val,
                    value1=val1,
                    value2=val2,
                )

            action_name = "обновлён" if chart_id else "создан"
            messages.success(
                request, f"График «{current_title}» успешно {action_name}!"
            )
            return redirect("forms_app:form17_view")

        elif action == "preview":
            table_data = parsed_data

    return render(
        request,
        "forms_app/form17.html",
        {
            "charts": charts,
            "table_data": table_data,
            "chart_data": _generate_chart_b64(
                table_data, current_title, current_label1, current_label2
            ),
            "current_title": current_title,
            "current_label1": current_label1,
            "current_label2": current_label2,
            "loaded_chart_id": loaded_chart_id,
        },
    )


@login_required
def form17_load_chart(request, pk):
    chart = get_object_or_404(ManualChart, pk=pk, user=request.user)
    data_points = chart.data_points.all()
    table_data = [(dp.date, dp.value1, dp.value2) for dp in data_points]

    charts = ManualChart.objects.filter(user=request.user).order_by("-created_at")

    return render(
        request,
        "forms_app/form17.html",
        {
            "charts": charts,
            "table_data": table_data,
            "chart_data": _generate_chart_b64(
                table_data, chart.title, chart.label1, chart.label2
            ),
            "current_title": chart.title,
            "current_label1": chart.label1,
            "current_label2": chart.label2,
            "loaded_chart_id": pk,
        },
    )


@login_required
def form17_delete_chart(request, pk):
    chart = get_object_or_404(ManualChart, pk=pk, user=request.user)
    title = chart.title
    chart.delete()
    messages.success(request, f"График «{title}» удалён.")
    return redirect("forms_app:form17_view")


def _generate_chart_b64(
    table_data, title="График", label1="Значение 1", label2="Значение 2"
):
    """
    Генерирует график в base64 с поддержкой двух осей и кастомных меток.
    table_data: список кортежей (date, value1, value2)
    """
    if not table_data:
        return None

    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import io
    import base64

    dates = [row[0] for row in table_data]
    values1 = [row[1] for row in table_data]
    values2 = [row[2] for row in table_data]

    fig, ax1 = plt.subplots(figsize=(10, 5))

    # --- Ось Y1: Значение 1 ---
    ax1.set_xlabel("Дата")
    ax1.set_ylabel(label1, color="tab:blue")
    line1 = ax1.plot(dates, values1, marker="o", color="tab:blue", label=label1)
    ax1.tick_params(axis="y", labelcolor="tab:blue")
    ax1.grid(True, linestyle="--", alpha=0.5)

    # --- Ось Y2: Значение 2 (если есть данные) ---
    has_value2 = any(v2 is not None for v2 in values2)
    if has_value2:
        ax2 = ax1.twinx()
        display_label2 = label2 if label2 else "Значение 2"
        ax2.set_ylabel(display_label2, color="tab:red")
        values2_clean = [v if v is not None else math.nan for v in values2]
        line2 = ax2.plot(
            dates, values2_clean, marker="s", color="tab:red", label=display_label2
        )
        ax2.tick_params(axis="y", labelcolor="tab:red")
    else:
        line2 = []

    # --- Легенда ---
    lines = line1 + line2
    labels = [l.get_label() for l in lines]
    if labels:
        ax1.legend(lines, labels, loc="upper left")

    plt.title(title)
    fig.tight_layout()
    plt.xticks(rotation=45)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    data = base64.b64encode(buf.read()).decode("utf-8")
    buf.close()
    return data
