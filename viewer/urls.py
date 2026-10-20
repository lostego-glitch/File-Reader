
from django.urls import path
from . import views, auth_views

urlpatterns = [
    path("", views.index, name="index"),
    path("login/", auth_views.login_view, name="login"),
    path("register/", auth_views.register_view, name="register"),
    path("logout/", auth_views.logout_view, name="logout"),
]
