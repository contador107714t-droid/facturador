FROM python:3.12-slim

WORKDIR /app

# Dependencias de sistema mínimas
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
  && rm -rf /var/lib/apt/lists/*

# Instala deps Python
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copia el código
COPY . /app

# Puerto interno
EXPOSE 8000

# Arranque (sin --reload en producción)
CMD ["python", "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
