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

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Форма 4: Накопительный отчёт"
        verbose_name_plural = "Форма 4: Накопительные отчёты"
        unique_together = ("user", "code", "date")  # Защита от дублей

    def __str__(self):
        return f"{self.user.username} — {self.code} — {self.date}"


# ------ Форма 8 -------


class Form8Report(models.Model):
    week_name = models.CharField("Неделя", max_length=100)
    date_extracted = models.DateField("Дата из файла", null=True, blank=True)

    # Суммы
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
        ordering = ["-uploaded_at"]
