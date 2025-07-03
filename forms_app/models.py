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
