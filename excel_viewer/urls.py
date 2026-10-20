
from django.contrib import admin
from django.urls import path
from viewer import views, auth_views

urlpatterns = [
    path("admin/", admin.site.urls),

    path("", views.index, name="index"),

    path("login/", auth_views.login_view, name="login"),

    path("accounts/login/", auth_views.login_view, name="accounts_login"),

    path("register/", auth_views.register_view, name="register"),

    path("logout/", auth_views.logout_view, name="logout"),
]
