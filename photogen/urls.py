from django.urls import path
from . import views

app_name = "photogen"

urlpatterns = [
    path("", views.generator, name="generator"),
    path("result/<int:task_id>/", views.result, name="result"),
    path("status/<int:task_id>/", views.task_status, name="status"),
    path("upload/", views.upload_asset, name="upload_asset"),
    path("library/", views.library, name="library"),
    path("library/delete/<int:asset_id>/", views.delete_asset, name="delete_asset"),
    path("history/", views.history, name="history"),
    path("task/<int:task_id>/", views.task_detail, name="task_detail"),
    path("task/delete/<int:task_id>/", views.delete_task, name="delete_task"),
    path("history/delete-bulk/", views.delete_tasks_bulk, name="delete_tasks_bulk"),
    # Промты (staff only)
    path("prompts/", views.prompt_list, name="prompt_list"),
    path("prompts/new/", views.prompt_edit, name="prompt_create"),
    path("prompts/<int:pk>/edit/", views.prompt_edit, name="prompt_edit"),
    path("prompts/<int:pk>/delete/", views.prompt_delete, name="prompt_delete"),
]