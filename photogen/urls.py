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
]