from django import forms
from .models import PromptPreset


class PromptPresetForm(forms.ModelForm):
    class Meta:
        model = PromptPreset
        fields = ["kind", "code", "title", "template", "order", "is_active"]
        widgets = {
            "template": forms.Textarea(attrs={
                "rows": 18,
                "style": "width:100%;font-family:monospace;",
            }),
            "title": forms.TextInput(attrs={"style": "width:100%;"}),
            "code": forms.TextInput(attrs={"style": "width:100%;"}),
        }
        help_texts = {
            "code": "Латиница, без пробелов. Для обёртки Muse — main.",
            "template": "Для «Обёртки Muse» используйте {pose} как плейсхолдер.",
        }