from django.contrib import admin
from django.urls import path, include
from create.views import generate_docx

urlpatterns = [
    path('admin/', admin.site.urls),
    path('generate-docx/', generate_docx, name='generate_docx'),
]
