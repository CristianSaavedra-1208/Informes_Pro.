FROM python:3.11-slim

WORKDIR /app

# Evita que Python escriba archivos .pyc y fuerza salida de logs
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Instalar dependencias del sistema operativo (necesarias para algunas librerías)
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el proyecto completo
COPY . .

# Exponer puerto default para Streamlit o FastAPI
EXPOSE 8501

# Comando por defecto (UI de Streamlit)
CMD ["streamlit", "run", "src/main.py", "--server.port=8501", "--server.address=0.0.0.0"]
