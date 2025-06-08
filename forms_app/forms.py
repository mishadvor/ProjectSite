# forms_app/forms.py

from django import forms


class UploadFileForm(forms.Form):
    file = forms.FileField(label="Загрузите отчет за неделю")
