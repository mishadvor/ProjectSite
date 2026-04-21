from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
from django.conf import settings


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
    clean_transfer_without_logistics = models.DecimalField(
        "Чистое Перечисление без Логистики",
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
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


# forms_app/models.py
from django.db import models
from django.contrib.auth.models import User


class Pattern15(models.Model):
    """Модель для хранения лекал Формы 15"""

    user = models.ForeignKey(
        User, on_delete=models.CASCADE, related_name="form15_patterns"
    )
    pattern_number = models.IntegerField(verbose_name="Номер лекала", default=0)
    name = models.CharField(max_length=100, verbose_name="Название лекала")
    width = models.IntegerField(verbose_name="Ширина (мм)")
    height = models.IntegerField(verbose_name="Высота (мм)")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Лекало (Форма 15)"
        verbose_name_plural = "Лекала (Форма 15)"
        ordering = ["pattern_number", "name"]

    def __str__(self):
        return (
            f"#{self.pattern_number:02d} - {self.name} ({self.width}×{self.height} мм)"
        )

    def save(self, *args, **kwargs):
        # Автоматически присваиваем номер при создании
        if not self.pk and self.pattern_number == 0:
            # Находим максимальный номер у пользователя
            max_number = Pattern15.objects.filter(user=self.user).aggregate(
                models.Max("pattern_number")
            )["pattern_number__max"]
            self.pattern_number = (max_number or 0) + 1
        super().save(*args, **kwargs)


from django.db import models
from django.contrib.auth.models import User
from django.core.validators import MinValueValidator, MaxValueValidator


class Form16Article(models.Model):
    """Модель для хранения 30 артикулов Формы 16"""

    user = models.ForeignKey(
        User, on_delete=models.CASCADE, verbose_name="Пользователь"
    )

    # ОДНО объявление поля position
    position = models.PositiveIntegerField(
        verbose_name="Позиция (1-50)",
        validators=[
            MinValueValidator(1, message="Позиция должна быть от 1 до 50"),
            MaxValueValidator(50, message="Позиция должна быть от 1 до 50"),
        ],
    )

    article_wb = models.CharField(max_length=100, verbose_name="Артикул WB")
    our_article = models.CharField(
        max_length=200,
        blank=True,
        verbose_name="Наш артикул",
        help_text="Наше внутреннее название артикула",
    )
    comments = models.TextField(
        blank=True,
        verbose_name="Комментарии",
        help_text="Заметки и комментарии по артикулу",
    )
    is_active = models.BooleanField(default=True, verbose_name="Активен")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Дата обновления")

    class Meta:
        verbose_name = "Артикул Формы 16"
        verbose_name_plural = "Артикулы Формы 16"
        unique_together = [
            "user",
            "position",
        ]  # Каждая позиция уникальна для пользователя
        ordering = ["position"]

    def __str__(self):
        return f"{self.position}. {self.article_wb} ({self.user.username})"


# forms_app/models.py
from django.db import models
from django.contrib.auth.models import User


class ManualChart(models.Model):
    title = models.CharField("Название графика", max_length=200)
    label1 = models.CharField(
        "Название Значения 1", max_length=100, default="Значение 1"
    )
    label2 = models.CharField(
        "Название Значения 2", max_length=100, default="Значение 2", blank=True
    )
    created_at = models.DateTimeField("Создано", auto_now_add=True)
    user = models.ForeignKey(User, on_delete=models.CASCADE)

    def __str__(self):
        return f"{self.title} ({self.user.username})"


class ManualChartDataPoint(models.Model):
    chart = models.ForeignKey(
        ManualChart, on_delete=models.CASCADE, related_name="data_points"
    )
    date = models.DateField("Дата")
    value1 = models.FloatField("Значение 1")
    value2 = models.FloatField(
        "Значение 2", null=True, blank=True
    )  # может отсутствовать

    class Meta:
        ordering = ["date"]
        unique_together = ("chart", "date")


# forms_app/models.py

from django.db import models
from django.contrib.auth.models import User


class ArticleCost(models.Model):
    user = models.ForeignKey(
        User, on_delete=models.CASCADE, verbose_name="Пользователь"
    )
    wb_article = models.CharField(
        max_length=50, verbose_name="WB артикул", db_index=True
    )
    seller_article = models.CharField(
        max_length=255, verbose_name="Артикул продавца", blank=True
    )
    cost = models.DecimalField(
        max_digits=12, decimal_places=2, verbose_name="Себестоимость"
    )

    class Meta:
        unique_together = ("user", "wb_article")
        verbose_name = "Себестоимость артикула"
        verbose_name_plural = "Себестоимости артикулов"

    def __str__(self):
        return f"{self.wb_article} → {self.cost} руб"


# Добавьте в forms_app/models.py


class Form20Data(models.Model):
    """Ежедневные данные (аналог Формы 4, но ежедневно)"""

    user = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.CASCADE, verbose_name="Пользователь"
    )
    code = models.CharField(
        max_length=50, db_index=True, verbose_name="Код номенклатуры"
    )
    article = models.CharField(
        max_length=200, blank=True, null=True, verbose_name="Артикул поставщика"
    )
    date = models.DateField(db_index=True, verbose_name="Дата отчета")

    # Финансовые показатели
    clear_sales_our = models.DecimalField(
        max_digits=15,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Чистые продажи Наши",
    )
    clear_sales_vb = models.DecimalField(
        max_digits=15,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Чистая реализация ВБ",
    )
    clear_transfer = models.DecimalField(
        max_digits=15,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Чистое Перечисление",
    )
    clear_transfer_without_log = models.DecimalField(
        max_digits=15,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Чистое Перечисление без Логистики",
    )

    # Средние цены
    our_price_mid = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Наша цена Средняя",
    )
    vb_selling_mid = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Реализация ВБ Средняя",
    )
    transfer_mid = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="К перечислению Среднее",
    )
    transfer_without_log_mid = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="К Перечислению без Логистики Средняя",
    )

    # Продажи и прибыль
    qentity_sale = models.IntegerField(
        null=True, blank=True, verbose_name="Чистые продажи, шт"
    )
    sebes_sale = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Себес Продаж (600р)",
    )
    profit_1 = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Прибыль на 1 Юбку",
    )
    percent_sell = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True, verbose_name="%Выкупа"
    )
    profit = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True, verbose_name="Прибыль"
    )
    orders = models.IntegerField(null=True, blank=True, verbose_name="Заказы")

    # Проценты
    percent_log_price = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="% Лог/Наша Цена",
    )
    spp_percent = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True, verbose_name="% СПП"
    )

    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Дата обновления")

    class Meta:
        verbose_name = "Форма 20 (Ежедневные данные)"
        verbose_name_plural = "Форма 20 (Ежедневные данные)"
        unique_together = [
            ["user", "code", "date"]
        ]  # Уникальность: пользователь + код + дата
        indexes = [
            models.Index(fields=["user", "code", "date"]),
            models.Index(fields=["date"]),
        ]
        ordering = ["-date", "code"]

    def __str__(self):
        return f"{self.code} - {self.date} ({self.user.username})"
