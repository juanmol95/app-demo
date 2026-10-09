# Imagen para publicar la app en Render (o cualquier servicio con Docker).
# Incluye Tesseract (lectura de la franja MRZ de la cédula).
FROM python:3.12-slim

RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr libglib2.0-0 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY servidor/requirements.txt servidor/requirements.txt
RUN pip install --no-cache-dir -r servidor/requirements.txt

COPY . .

# Render indica el puerto en la variable PORT; el servidor la usa.
ENV HOST=0.0.0.0
WORKDIR /app/servidor
CMD ["python", "App (1).py"]
