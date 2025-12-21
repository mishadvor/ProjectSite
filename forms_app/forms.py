# forms_app/forms.py

from django import forms
from .models import Form4Data, Form12Data  # Импортируем обе модели сразу


# Кастомный виджет для множественной загрузки файлов
# Этот класс определяется один раз и используется везде, где нужно multiple
class MultipleFileInput(forms.FileInput):
    allow_multiple_selected = True

    def __init__(self, attrs=None):
        # Устанавливаем атрибут multiple по умолчанию
        default_attrs = {"multiple": True}
        if attrs:
            default_attrs.update(attrs)
        super().__init__(attrs=default_attrs)


# --- Форма 4 ---


# Форма для загрузки одного файла (если нужна)
class UploadSingleFileForm(forms.Form):
    file = forms.FileField(label="Загрузите отчет за неделю")


# Форма для загрузки нескольких файлов (для Формы 4)
class UploadMultipleFileForm(forms.Form):
    file = forms.FileField(
        widget=MultipleFileInput(),
        label="Загрузите Excel-файлы",
        help_text="Поддерживаются .xlsx. Можно загружать несколько файлов.",
        required=True,  # или False, в зависимости от логики
    )


# Форма для редактирования данных (Форма 4)
class Form4DataForm(forms.ModelForm):
    class Meta:
        model = Form4Data
        fields = [
            "date",
            "article",
            "clear_sales_our",
            "clear_sales_vb",
            "clear_transfer",
            "clear_transfer_without_log",
            "our_price_mid",
            "vb_selling_mid",
            "transfer_mid",
            "transfer_without_log_mid",
            "qentity_sale",
            "sebes_sale",
            "profit_1",
            "percent_sell",
            "profit",
            "orders",
        ]
        widgets = {
            "date": forms.DateInput(attrs={"type": "date"}),
        }


# Универсальный алиас для Формы 4 (если используется)
UploadFileForm = UploadMultipleFileForm


# --- Форма 8 ---


class Form8UploadForm(forms.Form):
    files = forms.FileField(
        widget=MultipleFileInput(),
        label="Загрузите Excel-файлы",
        help_text="Поддерживается .xlsx. Можно загружать несколько файлов.",
        required=False,
    )


# --- Оборачиваемость ---


class ExcelProcessingForm(forms.Form):
    """Форма для обработки одного Excel-файла оборачиваемости"""

    file = forms.FileField(
        label="Загрузите отчёт по продажам/остаткам",
        help_text="Поддерживается .xlsx. Ожидается файл с данными за неделю.",
        widget=forms.FileInput(attrs={"accept": ".xlsx"}),
        required=True,
    )


# --- Форма 12 ---


class UploadFileForm12(forms.Form):
    file = forms.FileField(
        label="Загрузите Excel-файл (.xlsx)",
        widget=MultipleFileInput(),  # ✅ Используем кастомный виджет, поддерживающий multiple
        required=True,
    )


class Form12DataForm(forms.ModelForm):
    class Meta:
        model = Form12Data
        fields = "__all__"
        widgets = {
            "date": forms.DateInput(attrs={"type": "date"}),
        }


class UploadFileForm14(forms.Form):
    file = forms.FileField(
        label="Выберите файлы (.xlsx)"
        # НЕ указываем widget с multiple!
    )


from django import forms
from .models import Pattern15


class PatternForm(forms.ModelForm):
    """Форма для добавления/редактирования лекала"""

    class Meta:
        model = Pattern15
        fields = ["name", "width", "height"]
        widgets = {
            "name": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Например: Рукав, Спинка",
                }
            ),
            "width": forms.NumberInput(
                attrs={"class": "form-control", "placeholder": "в мм"}
            ),
            "height": forms.NumberInput(
                attrs={"class": "form-control", "placeholder": "в мм"}
            ),
        }


class CuttingForm(forms.Form):
    """Форма для параметров раскроя"""

    fabric_width = forms.IntegerField(
        label="Ширина полотна (мм)",
        initial=1500,
        min_value=100,
        max_value=5000,
        widget=forms.NumberInput(
            attrs={"class": "form-control", "placeholder": "1500"}
        ),
    )

    num_sets = forms.IntegerField(
        label="Количество комплектов",
        initial=1,
        min_value=1,
        max_value=100,
        widget=forms.NumberInput(attrs={"class": "form-control", "placeholder": "1"}),
    )

    output_format = forms.ChoiceField(
        label="Формат вывода",
        choices=[
            ("pdf", "PDF файл"),
            ("excel", "Excel с параметрами"),
        ],
        initial="pdf",
        widget=forms.RadioSelect(attrs={"class": "form-check-input"}),
    )


class Form16UploadForm(forms.Form):
    """Форма для загрузки файла оборачиваемости"""

    file = forms.FileField(
        label="Загрузите файл 'Детальная информация'",
        widget=forms.FileInput(attrs={"accept": ".xlsx"}),
        required=True,
    )


class Form16ArticleInputForm(forms.Form):
    """Форма для ввода 15 артикулов с сохранением порядка"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Создаем 15 полей с названиями от article_1 до article_15
        for i in range(1, 16):
            self.fields[f"article_{i}"] = forms.CharField(
                label=f"Артикул WB #{i}",
                required=(i <= 1),  # Только первый артикул обязателен
                widget=forms.TextInput(
                    attrs={"class": "form-control", "placeholder": "Введите артикул WB"}
                ),
                max_length=50,
            )


class Form16UploadForm(forms.Form):
    """Форма для загрузки файла оборачиваемости"""

    file = forms.FileField(
        label="Загрузите файл 'Детальная информация'",
        widget=forms.FileInput(attrs={"accept": ".xlsx"}),
        required=True,
        help_text="Ожидается файл с вкладкой 'Детальная информация'",
    )
