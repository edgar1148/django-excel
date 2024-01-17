from django.urls import path
from .views import upload_file, download_file

urlpatterns = [
    path('upload/', upload_file, name='upload_file'),
    path('download/<str:file_name>/', download_file, name='download_file'),
]
