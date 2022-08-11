from operator import mod
from pyexpat import model
from django.db import models

# Create your models here.

class Customer(models.Model):
    id = models.BigAutoField(primary_key=True, db_column='CustomerId')
    first_name = models.CharField(max_length=128, blank=True, db_column='FirstName')
    last_name = models.CharField(max_length=128, blank=True,db_column='LastName')
    address = models.CharField(max_length=512, blank=True, db_column='Address')

    def __str__(self):
        return f'{self.first_name}.{self.last_name}'

    def __unicode__(self):
        return 
