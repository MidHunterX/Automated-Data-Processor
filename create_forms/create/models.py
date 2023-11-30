from django.db import models


class Institution(models.Model):
    name = models.CharField(max_length=100)
    place = models.CharField(max_length=100, blank=True)
    district = models.CharField(max_length=100, blank=True)
    phone_no = models.CharField(max_length=100, blank=True)
    email = models.EmailField(blank=True)


class Student(models.Model):
    institution = models.ForeignKey(Institution, on_delete=models.CASCADE)
    student_name = models.CharField(max_length=50)
    student_class = models.CharField(max_length=50)
    student_ifsc = models.CharField(max_length=50)
    student_account = models.CharField(max_length=50)
    student_holder = models.CharField(max_length=50)
    student_branch = models.CharField(max_length=50)
