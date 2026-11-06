# ============================================================
# Stage 1: Build Next.js frontend
# ============================================================
FROM node:20-alpine AS frontend-builder

WORKDIR /app/frontend

# Copy package files first for layer caching
COPY frontend/package*.json ./
RUN npm ci --legacy-peer-deps

# Copy frontend source
COPY frontend/ .

# Build env — will be overridden at runtime via Render env vars
ARG NEXT_PUBLIC_API_URL=https://your-render-api.onrender.com
ENV NEXT_PUBLIC_API_URL=$NEXT_PUBLIC_API_URL

RUN npm run build

# ============================================================
# Stage 2: Python FastAPI backend + serve Next.js static files
# ============================================================
FROM python:3.11-slim AS backend

WORKDIR /app

# System dependencies for python-docx / lxml
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libxml2-dev \
    libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

# Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# FastAPI application source
COPY main.py .
COPY outputs/.gitkeep outputs/

# Copy Next.js standalone build (if using Next.js standalone output)
# For simplicity on Render, the frontend is deployed separately on Vercel
# or served statically — this Dockerfile is for the API backend only.

# Create non-root user for security
RUN useradd -m -u 1001 agentuser && chown -R agentuser:agentuser /app
USER agentuser

EXPOSE 8000

# Bind to 0.0.0.0 so Render can route traffic
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "2"]
