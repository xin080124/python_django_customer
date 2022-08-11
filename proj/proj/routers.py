from rest_framework import routers
from customers.viewsets import CustomerViewSet


router = routers.DefaultRouter()
router.register(r'customers', CustomerViewSet)
