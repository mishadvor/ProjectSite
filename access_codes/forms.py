from django import forms


class MasterCodeForm(forms.Form):
    code = forms.CharField(
        max_length=20,
        label="Код доступа",
        widget=forms.TextInput(
            attrs={
                "class": "form-control",
                "placeholder": "Введите мастер-код доступа",
                "autocomplete": "off",
            }
        ),
        error_messages={
            "required": "Пожалуйста, введите код доступа",
        },
    )
