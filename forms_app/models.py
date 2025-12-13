from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone


class UserReport(models.Model):
    REPORT_TYPES = (("form4", "Накопительный отчёт"), ("form5", "Остатки"))

    user = models.ForeignKey(User, on_delete=models.CASCADE)
    file_name = models.CharField(max_length=255)
    file_path = models.CharField(max_length=500)
    report_type = models.CharField(max_length=10, choices=REPORT_TYPES)
    last_updated = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.file_name} — {self.user.username}"


class StockRecord(models.Model):
    user = models.ForeignKey(
        User, on_delete=models.CASCADE, related_name="stock_records"
    )
    article_full_name = models.CharField("Артикул поставщика", max_length=255)
    size = models.CharField("Размер", max_length=10)
    quantity = models.IntegerField("Количество", default=0)
    location = models.CharField(
        "Место хранения", max_length=100, blank=True, null=True, default="Не указано"
    )
    note = models.TextField("Примечание", blank=True, null=True)

    def __str__(self):
        return f"{self.article_full_name} | {self.size} | {self.quantity}"

    class Meta:
        verbose_name = "Складская запись"
        verbose_name_plural = "Складские записи"


class WeeklyReport(models.Model):
    user = models.ForeignKey(
        User, on_delete=models.CASCADE, verbose_name="Пользователь"
    )
    week_name = models.CharField("Неделя", max_length=50)
    art_group = models.CharField("Группа артикулов", max_length=3)
    profit = models.FloatField("Прибыль")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.week_name} - {self.art_group}: {self.profit}"


# --- НОВАЯ МОДЕЛЬ ДЛЯ FORM4 (вместо Excel) ---
class Form4Data(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name="form4_data")
    code = models.CharField("Код номенклатуры", max_length=100)
    article = models.CharField("Артикул", max_length=100, blank=True, null=True)
    date = models.DateField("Дата")

    # Поля из Excel
    clear_sales_our = models.FloatField("Чистые продажи Наши", blank=True, null=True)
    clear_sales_vb = models.FloatField("Чистая реализация ВБ", blank=True, null=True)
    clear_transfer = models.FloatField("Чистое Перечисление", blank=True, null=True)
    clear_transfer_without_log = models.FloatField(
        "Чистое Перечисление без Логистики", blank=True, null=True
    )
    our_price_mid = models.FloatField("Наша цена Средняя", blank=True, null=True)
    vb_selling_mid = models.FloatField("Реализация ВБ Средняя", blank=True, null=True)
    transfer_mid = models.FloatField("К перечислению Среднее", blank=True, null=True)
    transfer_without_log_mid = models.FloatField(
        "К Перечислению без Логистики Средняя", blank=True, null=True
    )
    qentity_sale = models.IntegerField("Чистые продажи, шт", blank=True, null=True)
    sebes_sale = models.FloatField("Себес Продаж (600р)", blank=True, null=True)
    profit_1 = models.FloatField("Прибыль на 1 Юбку", blank=True, null=True)
    percent_sell = models.FloatField("%Выкупа", blank=True, null=True)
    profit = models.FloatField("Прибыль", blank=True, null=True)
    orders = models.IntegerField("Заказы", blank=True, null=True)
    percent_log_price = models.FloatField("% Лог/Наша Цена", blank=True, null=True)
    spp_percent = models.FloatField("% СПП", blank=True, null=True)  # <-- НОВОЕ ПОЛЕ

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Форма 4: Накопительный отчёт"
        verbose_name_plural = "Форма 4: Накопительные отчёты"
        unique_together = ("user", "code", "date")  # Защита от дублей

    def __str__(self):
        return f"{self.user.username} — {self.code} — {self.date}"


# ------ Форма 8 -------


class Form8Report(models.Model):
    user = models.ForeignKey(
        User, on_delete=models.CASCADE, verbose_name="Пользователь"
    )
    week_name = models.CharField("Неделя", max_length=100)
    date_extracted = models.DateField("Дата из файла", null=True, blank=True)

    # Суммыpython manage.py runserver
    profit = models.DecimalField(
        "Прибыль", max_digits=12, decimal_places=2, null=True, blank=True
    )
    clean_sales_ours = models.DecimalField(
        "Чистые продажи Наши", max_digits=12, decimal_places=2, null=True, blank=True
    )
    orders = models.IntegerField("Заказы", null=True, blank=True)

    # Средние (>0)
    spp_percent = models.DecimalField(
        "% СПП", max_digits=5, decimal_places=2, null=True, blank=True
    )
    avg_price = models.DecimalField(
        "Наша цена Средняя", max_digits=10, decimal_places=2, null=True, blank=True
    )
    profit_per_skirt = models.DecimalField(
        "Прибыль на 1 Юбку", max_digits=10, decimal_places=2, null=True, blank=True
    )
    pickup_rate = models.DecimalField(
        "% Выкупа", max_digits=5, decimal_places=2, null=True, blank=True
    )

    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.week_name

    class Meta:
        verbose_name = "Отчёт Формы 8"
        verbose_name_plural = "Форма 8 — Недельные метрики"
        # Опционально: запретить дубли у одного пользователя
        unique_together = ("user", "week_name")
        ordering = ["-uploaded_at"]


from datetime import date


class Form12Data(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    wb_article = models.CharField(max_length=100, verbose_name="Артикул WB")
    barcode = models.CharField(
        max_length=50, blank=True, null=True, verbose_name="Баркод"
    )
    seller_article = models.CharField(
        max_length=255, blank=True, null=True, verbose_name="Артикул продавца"
    )
    size = models.CharField(max_length=20, blank=True, null=True, verbose_name="Размер")
    orders_qty = models.IntegerField(blank=True, null=True, verbose_name="Заказы, шт.")
    order_amount_net = models.FloatField(
        blank=True, null=True, verbose_name="Сумма заказов минус комиссия WB, руб."
    )
    sold_qty = models.IntegerField(blank=True, null=True, verbose_name="Выкупили, шт.")
    transfer_amount = models.FloatField(
        blank=True, null=True, verbose_name="К перечислению за товар, руб."
    )
    current_stock = models.IntegerField(
        blank=True, null=True, verbose_name="Текущий остаток, шт."
    )
    date = models.DateField(verbose_name="Дата отчёта")

    class Meta:
        verbose_name = "Данные формы 12"
        verbose_name_plural = "Данные формы 12"
        ordering = ["-date"]

    def __str__(self):
        return f"{self.wb_article} ({self.date})"


class Form14Data(models.Model):
    """Форма 14: Агрегированные данные по всем артикулам (без разбивки по артикулам)"""

    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name="form14_data")
    date = models.DateField(verbose_name="Дата отчета")

    # Суммированные показатели за день
    total_orders_qty = models.IntegerField(
        verbose_name="Всего заказов, шт.", null=True, blank=True
    )
    total_order_amount_net = models.FloatField(
        verbose_name="Общая сумма заказов минус комиссия WB, руб.",
        null=True,
        blank=True,
    )
    total_sold_qty = models.IntegerField(
        verbose_name="Всего выкуплено, шт.", null=True, blank=True
    )
    total_transfer_amount = models.FloatField(
        verbose_name="Общая сумма к перечислению, руб.", null=True, blank=True
    )
    total_current_stock = models.IntegerField(
        verbose_name="Общий остаток, шт.", null=True, blank=True
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Данные формы 14"
        verbose_name_plural = "Данные формы 14"
        ordering = ["-date"]
        unique_together = ["user", "date"]  # Одна запись на пользователя за день

    def __str__(self):
        return f"Форма 14 - {self.date} ({self.user.username})"
