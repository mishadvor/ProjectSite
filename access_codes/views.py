from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required, user_passes_test
from django.contrib import messages
from django.conf import settings
from django.urls import reverse
from .forms import MasterCodeForm
from .models import MasterAccessCode


def is_admin(user):
    """Проверка, является ли пользователь администратором"""
    return user.is_superuser or user.is_staff


@login_required
def verify_code(request):
    """Страница ввода мастер-кода"""

    # Если уже верифицирован, перенаправляем
    if request.session.get(settings.TWO_FACTOR_AUTH["SESSION_KEY"], False):
        next_url = request.session.pop("next_url", reverse("forms_app:dashboard"))
        return redirect(next_url)

    if request.method == "POST":
        form = MasterCodeForm(request.POST)
        if form.is_valid():
            entered_code = form.cleaned_data["code"]

            # Проверяем мастер-код
            if MasterAccessCode.verify_code(entered_code):
                # Устанавливаем флаг в сессии
                request.session[settings.TWO_FACTOR_AUTH["SESSION_KEY"]] = True
                request.session.set_expiry(
                    settings.TWO_FACTOR_AUTH["CODE_EXPIRY_MINUTES"] * 60
                )

                messages.success(
                    request,
                    f"✅ Доступ подтверждён! Добро пожаловать, {request.user.username}.",
                )

                # Перенаправляем на сохранённый URL или дашборд
                next_url = request.session.pop(
                    "next_url", reverse("forms_app:dashboard")
                )
                return redirect(next_url)
            else:
                messages.error(request, "❌ Неверный код доступа!")
    else:
        form = MasterCodeForm()

    return render(request, "access_codes/verify_code.html", {"form": form})


@login_required
@user_passes_test(is_admin)
def manage_master_code(request):
    """Управление мастер-кодом (только для админов)"""

    # Получаем или создаём запись с мастер-кодом
    master_code, created = MasterAccessCode.objects.get_or_create(
        id=1,  # Одна запись с ID=1
        defaults={"code": "DEFAULT2024", "is_active": True},  # Код по умолчанию
    )

    if request.method == "POST":
        new_code = request.POST.get("code", "").strip().upper()
        is_active = request.POST.get("is_active") == "on"

        if new_code:
            master_code.code = new_code
            master_code.is_active = is_active
            master_code.updated_by = request.user
            master_code.save()

            # Очищаем кеш
            from django.core.cache import cache

            cache.delete("master_access_code")

            messages.success(request, f"✅ Мастер-код изменён на: {new_code}")
        else:
            messages.error(request, "❌ Код не может быть пустым")

        return redirect("access_codes:manage_master_code")

    return render(
        request,
        "access_codes/manage_master_code.html",
        {
            "master_code": master_code,
        },
    )
