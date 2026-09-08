FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
# libgomp1：onnxruntime（RapidOCR 依赖）在 slim 下需要 OpenMP 运行时，缺了 import 会失败
RUN pip install --no-cache-dir -r requirements.txt && \
    apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer fonts-noto-cjk fonts-liberation libgomp1 && \
    rm -rf /var/lib/apt/lists/*

COPY . .
RUN mkdir -p workspace/uploads workspace/output workspace/merged

EXPOSE 5000
CMD ["python", "review_app.py"]
