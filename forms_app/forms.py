# forms_app/forms.py
from django import forms
from .models import Form4Data


# Кастомный виджет для множественной загрузки
class MultipleFileInput(forms.FileInput):
    allow_multiple_selected = True


# Форма для загрузки одного файла (если нужна)
class UploadSingleFileForm(forms.Form):
    file = forms.FileField(label="Загрузите отчет за неделю")


# Форма для загрузки нескольких файлов
class UploadMultipleFileForm(forms.Form):
    file = forms.FileField(
        widget=MultipleFileInput(attrs={"multiple": True}),
        label="Загрузите Excel-файлы",
        help_text="Поддерживаются .xlsx. Можно загружать несколько файлов.",
        required=True,  # или False, в зависимости от логики
    )


# Универсальная форма (можно использовать вместо UploadMultipleFileForm)


class MultipleFileInput(forms.FileInput):
    allow_multiple_selected = True


class Form8UploadForm(forms.Form):
    files = forms.FileField(
        widget=MultipleFileInput(attrs={"multiple": True}),
        label="Загрузите Excel-файлы",
        help_text="Поддерживается .xlsx. Можно загружать несколько файлов.",
        required=False,
    )


# Форма для редактирования данных
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


# Добавьте в конец forms.py
UploadFileForm = UploadMultipleFileForm  # ← делаем алиас

# Оборачиваемость


class ExcelProcessingForm(forms.Form):
    """Форма для обработки одного Excel-файла оборачиваемости"""

    file = forms.FileField(
        label="Загрузите отчёт по продажам/остаткам",
        help_text="Поддерживается .xlsx. Ожидается файл с данными за неделю.",
        widget=forms.FileInput(attrs={"accept": ".xlsx"}),
        required=True,
    )
