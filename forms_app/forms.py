# forms_app/forms.py

from django import forms


class UploadFileForm(forms.Form):
    file = forms.FileField(label="Загрузите отчет за неделю")


class UploadExcelForm(forms.Form):
    excel_file = forms.FileField(label="Выберите Excel-файл")
