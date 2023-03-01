from django.urls import path
from . import views

urlpatterns = [
	path('', views.home),
    path('hva', views.hva),
    path('im', views.im),
    path('all', views.all),
    path('instellingsnaam', views.instellingsnaam),
    path('instellingscode', views.instellingscode),
    path('objectnummer', views.objectnummer),
    path('objectnmr', views.objectnmr),
    path('onderscheidendkenmerk', views.onderscheidendkenmerk),
    path('objectnaam', views.objectnaam),
    path('titel', views.titel),
    path('afbeelding', views.afbeelding),
    path('associatie', views.associatie),
    path('associatieplaats', views.associatieplaats),
    path('associatieperiode', views.associatieperiode),
    path('datum', views.datum),
    path('datumgroter', views.datumgroter),
    path('datumformat', views.datumformat),
    path('afmeting', views.afmeting),
    path('afmetingo', views.afmetingo),
    path('afmetingd', views.afmetingd),
    path('afmetingd', views.afmetingdd),
    path('rechten', views.rechten),
    path('rechtentype', views.rechtentype),
    path('rechtenref', views.rechtenref),
    path('pd', views.pd),
    path('toestand', views.toestand),
    path('verwerving', views.verwerving),
    path('variatitel', views.variatitel)
]