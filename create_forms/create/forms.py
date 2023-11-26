from django import forms
from .models import Institution, Student


class InstitutionForm(forms.ModelForm):
    class Meta:
        model = Institution
        fields = ["name", "place", "district", "phone_no", "email"]
        labels = {
            'name': 'Name of the Institution',
            'place': 'Place',
            'district': 'District',
            'phone_no': 'Phone number',
            'email': 'Email',
        }


class StudentForm(forms.ModelForm):
    class Meta:
        model = Student
        fields = ['student_name', 'student_class', 'student_ifsc']


StudentFormSet = forms.inlineformset_factory(Institution, Student, form=StudentForm, extra=12)
