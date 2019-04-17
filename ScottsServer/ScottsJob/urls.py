from django.urls import path
from django.conf.urls import url
from . import views

urlpatterns = [
    url(r'^$', views.HomePageView.as_view()),
    url(r'^AssignCharts/$', views.AssignChartView.as_view()),
]
