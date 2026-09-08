FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
# 系统库（slim 镜像缺失，否则运行时 import 失败）：
#   libgomp1           onnxruntime（RapidOCR 依赖）的 OpenMP 运行时
#   libgl1 libglib2.0-0 libsm6 libxrender1 libxext6   opencv(cv2, rapidocr 依赖) 的图形/字体库
RUN pip install --no-cache-dir -r requirements.txt && \
    apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer fonts-noto-cjk fonts-liberation \
        libgomp1 libgl1 libglib2.0-0 libsm6 libxrender1 libxext6 && \
    rm -rf /var/lib/apt/lists/*

COPY . .
RUN mkdir -p workspace/uploads workspace/output workspace/merged

EXPOSE 5000
CMD ["python", "review_app.py"]
