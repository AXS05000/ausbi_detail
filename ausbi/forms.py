from django import forms

class MesAnoForm(forms.Form):
    ano = forms.IntegerField()
    mes = forms.IntegerField()
