from django.db import models
from django.core.cache import cache


class MasterAccessCode(models.Model):
    """Модель для хранения мастер-кода доступа"""

    code = models.CharField(max_length=20, verbose_name="Код доступа")
    is_active = models.BooleanField(default=True, verbose_name="Активен")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Обновлён")
    updated_by = models.ForeignKey(
        "auth.User",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        verbose_name="Кем изменён",
    )

    class Meta:
        verbose_name = "Мастер-код доступа"
        verbose_name_plural = "Мастер-код доступа"

    def __str__(self):
        return (
            f"Мастер-код: {self.code} ({'Активен' if self.is_active else 'Неактивен'})"
        )

    @classmethod
    def get_active_code(cls):
        """Получить активный мастер-код (с кешированием)"""
        cache_key = "master_access_code"
        code_data = cache.get(cache_key)

        if code_data is None:
            try:
                master = cls.objects.get(is_active=True)
                code_data = master.code
                cache.set(cache_key, code_data, 300)  # Кеш на 5 минут
            except cls.DoesNotExist:
                code_data = None

        return code_data

    @classmethod
    def verify_code(cls, code):
        """Проверить введённый код"""
        active_code = cls.get_active_code()
        return active_code and code == active_code
