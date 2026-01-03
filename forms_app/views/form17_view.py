# forms_app/views/form17_view.py (обновлённая версия)

import json
from datetime import datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseBadRequest
from django.core.serializers.json import DjangoJSONEncoder
from forms_app.models import ManualChart, ManualChartDataPoint


@login_required
def form17_view(request):
    charts = (
        ManualChart.objects.filter(user=request.user)
        .prefetch_related("data_points")
        .order_by("-created_at")
    )
    table_data = []
    current_title = ""
    current_label1 = "Значение 1"
    current_label2 = ""
    loaded_chart_id = None
    chart_js_data = None

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
                    "chart_js_data": None,
                },
            )

        if not cleaned:
            messages.error(request, "Нет данных для сохранения или отображения.")
            return render(
                request,
                "forms_app/form17.html",
                {
                    "charts": charts,
                    "table_data": [],  # ← исправлено: было table_build
                    "current_title": current_title,
                    "current_label1": current_label1,
                    "current_label2": current_label2,
                    "loaded_chart_id": loaded_chart_id,
                    "chart_js_data": None,
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
                    "chart_js_data": None,
                },
            )

        # Подготовка данных для Chart.js
        labels = [d.strftime("%d.%m.%Y") for d, _, _ in parsed_data]
        values1 = [v1 for _, v1, _ in parsed_data]
        values2 = [v2 if v2 is not None else None for _, _, v2 in parsed_data]
        chart_js_data = {
            "labels": labels,
            "datasets": [
                {
                    "label": current_label1,
                    "data": values1,
                    "borderColor": "rgb(54, 162, 235)",
                    "backgroundColor": "rgba(54, 162, 235, 0.1)",
                    "fill": False,
                    "tension": 0,
                }
            ],
        }
        if any(v2 is not None for v2 in values2):
            chart_js_data["datasets"].append(
                {
                    "label": current_label2 or "Значение 2",
                    "data": values2,
                    "borderColor": "rgb(255, 99, 132)",
                    "backgroundColor": "rgba(255, 99, 132, 0.1)",
                    "fill": False,
                    "tension": 0,
                }
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
            "current_title": current_title,
            "current_label1": current_label1,
            "current_label2": current_label2,
            "loaded_chart_id": loaded_chart_id,
            "chart_js_data": (
                json.dumps(chart_js_data, cls=DjangoJSONEncoder)
                if chart_js_data
                else None
            ),
        },
    )


@login_required
def form17_load_chart(request, pk):
    chart = get_object_or_404(ManualChart, pk=pk, user=request.user)
    data_points = chart.data_points.all().order_by("date")
    table_data = [(dp.date, dp.value1, dp.value2) for dp in data_points]

    charts = ManualChart.objects.filter(user=request.user).order_by("-created_at")

    # Подготовка данных для Chart.js
    labels = [dp.date.strftime("%d.%m.%Y") for dp in data_points]
    values1 = [dp.value1 for dp in data_points]
    values2 = [dp.value2 for dp in data_points]

    chart_js_data = {
        "labels": labels,
        "datasets": [
            {
                "label": chart.label1,
                "data": values1,
                "borderColor": "rgb(54, 162, 235)",
                "backgroundColor": "rgba(54, 162, 235, 0.1)",
                "fill": False,
                "tension": 0,
            }
        ],
    }
    if any(v2 is not None for v2 in values2):
        chart_js_data["datasets"].append(
            {
                "label": chart.label2 or "Значение 2",
                "data": values2,
                "borderColor": "rgb(255, 99, 132)",
                "backgroundColor": "rgba(255, 99, 132, 0.1)",
                "fill": False,
                "tension": 0,
            }
        )

    return render(
        request,
        "forms_app/form17.html",
        {
            "charts": charts,
            "table_data": table_data,
            "current_title": chart.title,
            "current_label1": chart.label1,
            "current_label2": chart.label2,
            "loaded_chart_id": pk,
            "chart_js_data": json.dumps(chart_js_data, cls=DjangoJSONEncoder),
        },
    )


@login_required
def form17_delete_chart(request, pk):
    chart = get_object_or_404(ManualChart, pk=pk, user=request.user)
    title = chart.title
    chart.delete()
    messages.success(request, f"График «{title}» удалён.")
    return redirect("forms_app:form17_view")
