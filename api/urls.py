from django.urls import path
from .views import ProcessExcelView

urlpatterns = [
    path('process/', ProcessExcelView.as_view(), name='process_excel'),
]
