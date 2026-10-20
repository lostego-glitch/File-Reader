
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.forms import UserCreationForm
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST


def login_view(request):
    if request.user.is_authenticated:
        return redirect("index")

    next_url = request.POST.get("next") or request.GET.get("next", "")

    if request.method == "POST":
        username = request.POST.get("username", "").strip()
        password = request.POST.get("password", "")

        user = authenticate(
            request,
            username=username,
            password=password
        )

        if user is not None:
            login(request, user)

            if next_url.startswith("/") and not next_url.startswith("//"):
                return redirect(next_url)

            return redirect("index")

        messages.error(request, "Invalid username or password.")

    return render(
        request,
        "viewer/login.html",
        {"next": next_url}
    )


def register_view(request):
    if request.user.is_authenticated:
        return redirect("index")

    form = UserCreationForm(request.POST or None)

    if request.method == "POST" and form.is_valid():
        user = form.save()
        login(request, user)
        messages.success(request, "Your account was created successfully.")
        return redirect("index")

    return render(
        request,
        "viewer/register.html",
        {"form": form}
    )


@require_POST
def logout_view(request):
    logout(request)
    messages.success(request, "You have logged out successfully.")
    return redirect("login")
