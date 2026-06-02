FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY gestionale/ .

RUN mkdir -p /app/data

EXPOSE 8000

CMD ["sh", "-c", "python manage.py migrate && python manage.py shell -c \"from django.contrib.auth.models import User; User.objects.filter(username='cocca').exists() or User.objects.create_superuser('cocca', '', 'Villorba26!')\" && python manage.py runserver 0.0.0.0:8000"]
