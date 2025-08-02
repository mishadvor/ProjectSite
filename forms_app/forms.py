# forms_app/forms.py

from django import forms
from .models import Form4Data


class UploadFileForm(forms.Form):
    file = forms.FileField(label="Загрузите отчет за неделю")


class UploadExcelForm(forms.Form):
    excel_file = forms.FileField(label="Выберите Excel-файл")


# forms_app/forms.py


class UploadFileForm(forms.Form):
    file = forms.FileField(
        label="Выберите файл", widget=forms.FileInput(attrs={"accept": ".xlsx,.xls"})
    )


class Form4DataForm(forms.ModelForm):
    """
    Форма для редактирования одной записи Form4Data
    """

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


# Кастомный виджет, который поддерживает множественную загрузку
class MultipleFileInput(forms.FileInput):
    allow_multiple_selected = True


class Form8UploadForm(forms.Form):
    files = forms.FileField(
        widget=MultipleFileInput(attrs={"multiple": True}),
        required=False,  # 🔑 Отключаем валидацию на "обязательность"
        label="Загрузите Excel-файлы",
        help_text="Поддерживаются .xlsx. Можно загружать несколько файлов.",
    )
