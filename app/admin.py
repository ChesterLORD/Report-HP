from django.contrib import admin

from .models import  HospitalMRD, ClinicalPharmacist, AddNote


admin.site.register(HospitalMRD)
admin.site.register(ClinicalPharmacist)
admin.site.register(AddNote)


