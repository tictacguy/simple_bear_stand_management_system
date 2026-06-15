from django.db import models
from django.contrib.auth.models import User


class Category(models.Model):
    name = models.CharField(max_length=50, unique=True)

    class Meta:
        verbose_name_plural = "categories"

    def __str__(self):
        return self.name


class Product(models.Model):
    ICON_CHOICES = [
        ("coffee", "Caffè"), ("cup-soda", "Bibita"), ("beer", "Birra"), ("wine", "Vino"),
        ("martini", "Cocktail"), ("glass-water", "Acqua"), ("milk", "Latte"),
        ("sandwich", "Panino"), ("pizza", "Pizza"), ("salad", "Insalata"),
        ("croissant", "Cornetto"), ("cake-slice", "Dolce"), ("cookie", "Biscotto"),
        ("ice-cream-cone", "Gelato"), ("popcorn", "Snack"), ("egg-fried", "Uovo"),
        ("beef", "Carne"), ("fish", "Pesce"), ("apple", "Frutta"), ("carrot", "Verdura"),
        ("utensils", "Piatto"), ("soup", "Zuppa"), ("candy", "Caramella"), ("wheat", "Pane"),
    ]

    PRINT_DEST_CHOICES = [
        ("cucina", "Cucina"),
        ("bar", "Bar"),
    ]

    name = models.CharField(max_length=100)
    category = models.ForeignKey(Category, on_delete=models.SET_NULL, null=True, blank=True)
    price = models.DecimalField(max_digits=7, decimal_places=2)
    stock = models.PositiveIntegerField(null=True, blank=True)
    icon = models.CharField(max_length=30, blank=True, default="", choices=ICON_CHOICES)
    is_shortcut = models.BooleanField(default=False, help_text="Mostra come tab rapida in cassa")
    print_destinations = models.CharField(max_length=50, blank=True, default="", help_text="Destinazioni extra: cucina,bar")

    def __str__(self):
        return f"{self.name} - €{self.price}"


class Order(models.Model):
    PAYMENT_CHOICES = [("cash", "Contanti"), ("card", "Carta")]

    created_at = models.DateTimeField(auto_now_add=True)
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    payment_method = models.CharField(max_length=4, choices=PAYMENT_CHOICES)
    discount_percent = models.DecimalField(max_digits=5, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=9, decimal_places=2, default=0)
    note = models.CharField(max_length=300, blank=True, default="")

    def __str__(self):
        return f"Ordine #{self.pk} - €{self.total}"


class OrderItem(models.Model):
    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name="items")
    product = models.ForeignKey(Product, on_delete=models.SET_NULL, null=True)
    product_name = models.CharField(max_length=100)
    price = models.DecimalField(max_digits=7, decimal_places=2)
    quantity = models.PositiveIntegerField(default=1)

    @property
    def subtotal(self):
        return self.price * self.quantity


class Operator(models.Model):
    name = models.CharField(max_length=100, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name


class CashWithdrawal(models.Model):
    operator = models.ForeignKey(Operator, on_delete=models.CASCADE, related_name="withdrawals")
    amount = models.DecimalField(max_digits=9, decimal_places=2)
    note = models.CharField(max_length=200, blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.operator.name} - €{self.amount}"
