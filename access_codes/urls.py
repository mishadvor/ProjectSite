from django.urls import path
from . import views

app_name = "access_codes"

urlpatterns = [
    path("verify/", views.verify_code, name="verify_code"),
    path("manage-master/", views.manage_master_code, name="manage_master_code"),
]
