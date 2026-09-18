# Build the React single-page app, then serve it from the FastAPI service.
FROM node:22-bookworm-slim AS frontend-build

WORKDIR /build/frontend

COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci

COPY frontend/ ./
RUN npm run build


FROM python:3.13-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# antiword enables the optional legacy .doc import supported by the API.
RUN apt-get update \
    && apt-get install --no-install-recommends -y antiword \
    && rm -rf /var/lib/apt/lists/*

COPY backend/requirements.txt ./backend/requirements.txt
RUN pip install --no-cache-dir -r ./backend/requirements.txt

COPY backend/ ./backend/
COPY --from=frontend-build /build/frontend/dist ./frontend/dist

EXPOSE 8000

CMD ["sh", "-c", "uvicorn app:app --app-dir /app/backend --host 0.0.0.0 --port ${PORT:-8000}"]
