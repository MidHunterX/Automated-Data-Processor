from django.db import models


class Institution(models.Model):
    name = models.CharField(max_length=100)
    place = models.CharField(max_length=100)
    district = models.CharField(max_length=100)
    phone_no = models.CharField(max_length=100)
    email = models.EmailField()


class Student(models.Model):
    institution = models.ForeignKey(Institution, on_delete=models.CASCADE)
    student_name = models.CharField(max_length=100)
    student_class = models.CharField(max_length=50)
    student_ifsc = models.CharField(max_length=20)
