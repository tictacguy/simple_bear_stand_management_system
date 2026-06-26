from django.urls import path
from . import views

app_name = "cassa"

urlpatterns = [
    path("", views.cassa, name="cassa"),
    path("search/", views.search_products, name="search"),
    path("order/", views.create_order, name="create_order"),
    path("order/delete/<int:pk>/", views.delete_order, name="delete_order"),
    path("order/print/<int:pk>/", views.print_order, name="print_order"),
    path("ordini/", views.ordini, name="ordini"),
    path("withdrawal/", views.create_withdrawal, name="create_withdrawal"),
    path("inventario/", views.inventario, name="inventario"),
    path("inventario/save/", views.product_save, name="product_save"),
    path("inventario/delete/<int:pk>/", views.product_delete, name="product_delete"),
    path("operatori/", views.operatori, name="operatori"),
    path("operatori/save/", views.operator_save, name="operator_save"),
    path("operatori/delete/<int:pk>/", views.operator_delete, name="operator_delete"),
    path("andamento/", views.andamento, name="andamento"),
    path("andamento/data/", views.andamento_data, name="andamento_data"),
    path("resoconti/", views.resoconti, name="resoconti"),
    path("resoconti/download/<str:date>/", views.resoconti_download, name="resoconti_download"),
    path("resoconti/export-all/", views.resoconti_export_all, name="resoconti_export_all"),
    path("impostazioni/", views.impostazioni, name="impostazioni"),
    path("impostazioni/printer/save/", views.printer_save, name="printer_save"),
    path("impostazioni/printer/delete/<int:pk>/", views.printer_delete, name="printer_delete"),
    path("impostazioni/printer/test/<int:pk>/", views.printer_test, name="printer_test"),
    path("impostazioni/scan/", views.network_scan, name="network_scan"),
]
