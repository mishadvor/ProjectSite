import pandas as pd
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows


def form1(request):
    if request.method == "POST":
        mode = request.POST.get("mode")
        start_date_str = request.POST.get("start_date")

        try:
            start_date = pd.to_datetime(start_date_str)
        except Exception:
            return render(
                request, "forms_app/form1.html", {"error": "Некорректная дата."}
            )

        dfs = []

        if mode == "single":
            file = request.FILES.get("file_single")
            if not file:
                return render(
                    request,
                    "forms_app/form1.html",
                    {"error": "Необходимо загрузить файл."},
                )
            df = pd.read_excel(file)
            df["Источник"] = "Точка 1"
            dfs.append(df)

        elif mode == "multiple":
            file1 = request.FILES.get("file1")
            file2 = request.FILES.get("file2")
            file3 = request.FILES.get("file3")

            if not (file1 and file2 and file3):
                return render(
                    request,
                    "forms_app/form1.html",
                    {"error": "Пожалуйста, загрузите все три файла."},
                )

            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
            df3 = pd.read_excel(file3)

            df1["Источник"] = "Точка 1"
            df2["Источник"] = "Точка 2"
            df3["Источник"] = "Точка 3"

            dfs = [df1, df2, df3]

        # Объединяем данные
        combined_df = pd.concat(dfs, ignore_index=True)

        # Чистка дат
        combined_df["Дата конца"] = (
            combined_df["Дата конца"].astype(str).str.split("T").str[0]
        )
        combined_df["Дата конца"] = pd.to_datetime(
            combined_df["Дата конца"], errors="coerce"
        )
        combined_df["Дата конца"] = combined_df["Дата конца"].dt.strftime("%Y-%m-%d")
        combined_df["Дата конца"] = pd.to_datetime(combined_df["Дата конца"])

        # Фильтрация по дате
        filtered_df = combined_df[combined_df["Дата конца"] >= start_date]

        # Агрегация данных
        sums_per_date = (
            filtered_df.groupby("Дата конца")
            .agg(
                {
                    "Продажа": "sum",
                    "К перечислению за товар": "sum",
                    "Стоимость логистики": "sum",
                    "Общая сумма штрафов": "sum",
                    "Стоимость хранения": "sum",
                    "Стоимость операций на приемке": "sum",
                    "Прочие удержания/выплаты": "sum",
                    "Итого к оплате": "sum",
                }
            )
            .astype(int)
            .reset_index()
        )

        sums_per_date["Дата конца"] = sums_per_date["Дата конца"].dt.strftime(
            "%d-%m-%Y"
        )

        # Добавляем столбец "Наш %"
        sums_per_date["Наш %"] = (
            sums_per_date["Итого к оплате"] / sums_per_date["Продажа"] * 100
        ).round(1)

        # Построение графика для основных финансовых показателей
        buf_main = BytesIO()
        plt.figure(figsize=(15, 8))
        for column in [
            "Продажа",
            "К перечислению за товар",
            "Стоимость логистики",
            "Итого к оплате",
        ]:
            plt.plot(
                sums_per_date["Дата конца"],
                sums_per_date[column],
                label=column,
                marker="o",  # Добавляем точки
                markersize=6,  # Размер точек
            )
        plt.title(
            f'Основные финансовые показатели (с {start_date.strftime("%d-%m-%Y")})'
        )
        plt.xlabel("Дата")
        plt.ylabel("Сумма")
        plt.xticks(rotation=90)
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(buf_main, format="png")
        plt.close()

        # Построение отдельного графика для "Наш%"
        buf_nash_percent = BytesIO()
        plt.figure(figsize=(15, 4))
        plt.plot(
            sums_per_date["Дата конца"],
            sums_per_date["Наш %"],
            label="Наш %",
            color="red",
            marker="o",  # Добавляем точки
            markersize=4,  # Размер точек
        )
        plt.title(f'Наш Процент (с {start_date.strftime("%d-%m-%Y")})')
        plt.xlabel("Дата")
        plt.ylabel("Наш Процент")
        plt.xticks(rotation=90)
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(buf_nash_percent, format="png")
        plt.close()

        # Создание Excel-файла
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            workbook = writer.book
            worksheet = workbook.create_sheet(title="Report")

            # Добавляем таблицу
            for row in dataframe_to_rows(sums_per_date, index=False, header=True):
                worksheet.append(row)

            # Добавляем график для основных финансовых показателей
            img_main = OpenpyxlImage(buf_main)
            worksheet.add_image(img_main, "K10")

            # Добавляем отдельный график для "Наш%"
            img_nash_percent = OpenpyxlImage(buf_nash_percent)
            worksheet.add_image(
                img_nash_percent, "K50"
            )  # Вы можете изменить позицию как удобно

        output.seek(0)

        # Возвращаем файл
        response = HttpResponse(
            output.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = (
            f'attachment; filename=wildberries_report_Form_1_{start_date.strftime("%Y%m%d")}.xlsx'
        )
        return response

    return render(request, "forms_app/form1.html")
