from django.contrib import admin
from .models import Category, Product, Order, OrderItem


class OrderItemInline(admin.TabularInline):
    model = OrderItem
    extra = 0


@admin.register(Category)
class CategoryAdmin(admin.ModelAdmin):
    list_display = ["name"]


@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ["name", "category", "price", "stock", "is_shortcut"]
    list_filter = ["category", "is_shortcut"]


@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display = ["id", "created_at", "user", "payment_method", "total"]
    list_filter = ["payment_method", "created_at"]
    inlines = [OrderItemInline]
