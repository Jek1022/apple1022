from django.conf.urls import url
from . import views

urlpatterns = [
    url(r'^$', views.IndexView.as_view(), name='index'),
    url(r'^pdf/$', views.Pdf.as_view()),
    url(r'^(?P<pk>[0-9]+)/$', views.DetailView.as_view(), name='detail'),
    url(r'^create/$', views.CreateView.as_view(), name='create'),
    url(r'^(?P<pk>[0-9]+)/update/$', views.UpdateView.as_view(), name='update'),
    url(r'^(?P<pk>[0-9]+)/delete/$', views.DeleteView.as_view(), name='delete'),
    url(r'^(?P<pk>[0-9]+)/pdf/$', views.Pdf.as_view(), name='pdf'),
    url(r'^approve/$', views.approve, name='approve'),
    url(r'^disapprove/$', views.disapprove, name='disapprove'),
    url(r'^posting/$', views.posting, name='posting'),
    url(r'^release/$', views.release, name='release'),
    url(r'^importreprfv/$', views.importreprfv, name='importreprfv'),
    url(r'^importreppcv/$', views.importreppcv, name='importreppcv'),
    url(r'^report/$', views.ReportView.as_view(), name='report'),
    url(r'^pdf2/$', views.GeneratePDF.as_view(), name='pdf2'),
    url(r'^excel/$', views.GenerateExcel.as_view(), name='excel'),
    url(r'^generatedefaultentries/$', views.generatedefaultentries, name='generatedefaultentries'),
    url(r'^searchforposting/$', views.searchforposting, name='searchforposting'),
    url(r'^searchfordigibanker/$', views.searchfordigibanker, name='searchfordigibanker'),
    url(r'^gopost/$', views.gopost, name='gopost'),
    url(r'^gounpost/$', views.gounpost, name='gounpost'),
    url(r'^digibanker/$', views.digibanker, name='digibanker'),
    url(r'^upload/$', views.upload, name='upload'),
    #url(r'^(?P<pk>[0-9]+)/disapproved/$', views.DisapprovedView.as_view(), name='disapproved'),
]
