from django.shortcuts import redirect
from django.urls import reverse
from django.conf import settings


class TwoFactorMiddleware:
    """Middleware для проверки двухфакторной аутентификации"""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        # Список URL, которые не требуют 2FA
        exempt_urls = [
            reverse("access_codes:verify_code"),
            reverse("accounts:logout"),
            "/admin/",
            "/accounts/login/",
        ]

        # Проверяем, нужна ли 2FA
        if request.user.is_authenticated:
            current_path = request.path_info

            # Если 2FA уже подтверждена
            if request.session.get(settings.TWO_FACTOR_AUTH["SESSION_KEY"], False):
                return self.get_response(request)

            # Проверяем, не находится ли пользователь на exempt URL
            is_exempt = any(current_path.startswith(url) for url in exempt_urls)

            if not is_exempt:
                # Сохраняем текущий URL для редиректа после верификации
                request.session["next_url"] = request.get_full_path()
                return redirect("access_codes:verify_code")

        return self.get_response(request)
