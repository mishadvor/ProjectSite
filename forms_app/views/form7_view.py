import pandas as pd
import numpy as np
import plotly.express as px
import json
from plotly.utils import PlotlyJSONEncoder
from django.shortcuts import render
from django.core.files.storage import default_storage
from ..models import WeeklyReport
from django.shortcuts import redirect
from django.contrib import messages
from io import BytesIO
from django.contrib.auth.decorators import login_required


# Функция извлечения первых 3 цифр артикула
def get_art_prefix(art):
    try:
        digits = "".join([c for c in str(art).strip() if c.isdigit()])
        return digits[:3] if len(digits) >= 3 else "Unknown"
    except:
        return "Unknown"


# Форма загрузки файла
@login_required
def form7_upload(request):
    if request.method == "POST":
        excel_file = request.FILES.get("excel_file")
        if not excel_file:
            return render(
                request,
                "forms_app/form7/upload.html",
                {"error": "Выберите файл"},
            )

        try:
            # Сохраняем временно файл
            file_path = default_storage.save(excel_file.name, excel_file)
            full_path = default_storage.path(file_path)

            # Чтение Excel
            xls = pd.ExcelFile(full_path)
            if "Основные данные" not in xls.sheet_names:
                return render(
                    request,
                    "forms_app/form7/upload.html",
                    {"error": 'Лист "Основные данные" не найден'},
                )

            df = pd.read_excel(xls, sheet_name="Основные данные")

            # Поиск нужных колонок
            art_col = [
                col
                for col in df.columns
                if "артикул" in str(col).lower() and "поставщика" in str(col).lower()
            ]
            profit_col = [col for col in df.columns if "прибыль" in str(col).lower()]

            if not art_col or not profit_col:
                return render(
                    request,
                    "forms_app/form7/upload.html",
                    {"error": "Не найдены нужные колонки"},
                )

            art_col = art_col[0]
            profit_col = profit_col[0]

            # Очистка данных
            df = df.dropna(subset=[art_col, profit_col])
            df[art_col] = df[art_col].astype(str).str.strip()
            df[profit_col] = pd.to_numeric(df[profit_col], errors="coerce")

            week_name = excel_file.name.replace(".xlsx", "").strip()

            # Извлечение групп артикулов
            df["группа"] = df[art_col].apply(get_art_prefix)
            grouped = df.groupby("группа")[profit_col].sum().reset_index()

            # Сохранение в БД
            for _, row in grouped.iterrows():
                WeeklyReport.objects.update_or_create(
                    user=request.user,
                    week_name=week_name,
                    art_group=row["группа"],
                    defaults={"profit": row[profit_col]},
                )

            return render(
                request,
                "forms_app/form7/success.html",
                {
                    "week": week_name,
                    "groups": grouped["группа"].tolist(),
                    "total_profit": round(grouped[profit_col].sum(), 2),
                },
            )

        except Exception as e:
            return render(
                request,
                "forms_app/form7/upload.html",
                {"error": f"Ошибка: {e}"},
            )

    return render(request, "forms_app/form7/upload.html")


# Функция отрисовки графика


@login_required
def form7_graph(request):
    try:
        # Получаем данные
        reports = WeeklyReport.objects.filter(user=request.user).values(
            "week_name", "art_group", "profit"
        )  # ✅ Только свои
        if not reports:
            return render(request, "forms_app/form7/no_data.html")

        df = pd.DataFrame(reports)

        # Очистка данных (как в Colab)
        df = df.dropna()
        df = df[df["profit"] != 0]  # Удаляем нулевые значения

        if df.empty:
            return render(request, "forms_app/form7/no_data.html")

        # Подготовка данных
        df["size"] = df["profit"].abs()
        df["type"] = np.where(df["profit"] >= 0, "Прибыль", "Убыток")

        # Создаем график через словарь (как в Plotly.js)
        graph_data = {
            "data": [
                {
                    "x": df["art_group"].tolist(),
                    "y": df["week_name"].tolist(),
                    "z": df["profit"].tolist(),
                    "mode": "markers",
                    "type": "scatter3d",
                    "marker": {
                        "size": (df["size"] / 10).tolist(),  # Масштабирование
                        "sizeref": 0.3,
                        "color": df["type"]
                        .map({"Прибыль": "green", "Убыток": "red"})
                        .tolist(),
                        "sizemode": "area",
                        "opacity": 0.8,
                    },
                    "name": "Прибыль/Убыток",
                    "hovertemplate": "Группа: %{x}<br>Период: %{y}<br>Прибыль: %{z:.1f} руб.",
                }
            ],
            "layout": {
                "title": "Динамика прибыли по группам артикулов",
                "scene": {
                    "xaxis": {"title": "Группа артикула"},
                    "yaxis": {"title": "Период"},
                    "zaxis": {"title": "Прибыль (руб)"},
                },
                "height": 800,
                "showlegend": True,
            },
        }

        return render(
            request,
            "forms_app/form7/graph.html",
            {
                "graph_json": json.dumps(graph_data),
                "debug_info": f"Данные: {len(df)} строк",
            },
        )

    except Exception as e:
        print(f"Error: {str(e)}")
        return render(
            request,
            "forms_app/form7/error.html",
            {"error": f"Ошибка при построении графика: {str(e)}"},
        )


@login_required
def clear_form7_data(request):
    try:
        # Получаем количество записей перед удалением
        count = WeeklyReport.objects.count()

        # Удаляем все данные
        WeeklyReport.objects.filter(user=request.user).delete()  # ✅ Только свои

        # Сообщение об успехе
        messages.success(
            request,
            f"Успешно удалено {count} записей. Теперь можно загрузить новые данные.",
        )

    except Exception as e:
        messages.error(request, f"Ошибка при очистке данных: {str(e)}")

    return redirect("forms_app:form7_upload")  # Перенаправляем на страницу загрузки
