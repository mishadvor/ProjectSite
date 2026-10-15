import json
import requests
from django.conf import settings
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt

AITUNNEL_URL = "https://api.aitunnel.ru/v1/chat/completions"

@csrf_exempt
def chat_with_ai(request):
    if request.method != "POST":
        return JsonResponse({"error": "Метод не поддерживается"}, status=405)

    try:
        data = json.loads(request.body)
        user_message = (data.get("message") or "").strip()
        if not user_message:
            return JsonResponse({"error": "Пустое сообщение"}, status=400)

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {settings.AITUNNEL_API_KEY}",
        }
        payload = {
            "model": "deepseek-v4-pro",   # замените на нужную модель
            "messages": [
                {"role": "user", "content": user_message},
            ],
        }

        response = requests.post(AITUNNEL_URL, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        answer = result["choices"][0]["message"]["content"]
        return JsonResponse({"answer": answer})

    except requests.exceptions.Timeout:
        return JsonResponse({"error": "Таймаут от AI"}, status=504)
    except requests.exceptions.RequestException as e:
        return JsonResponse({"error": f"Ошибка запроса к AI: {e}"}, status=502)
    except (KeyError, IndexError) as e:
        return JsonResponse({"error": f"Не удалось разобрать ответ: {e}"}, status=500)
