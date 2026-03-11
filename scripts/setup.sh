#!/bin/bash
set -e

echo "Setting up Excel AI..."

cp .env.example .env

docker-compose up -d postgres
echo "Waiting for PostgreSQL..."
sleep 5

docker-compose up -d ollama
echo "Pulling LLM model (first run: 5-10 min)..."
docker exec $(docker-compose ps -q ollama) ollama pull llama3.1:8b

docker-compose up -d backend

echo ""
echo "Excel AI is running!"
echo "  API:    http://localhost:8000"
echo "  Docs:   http://localhost:8000/docs"
echo "  Health: http://localhost:8000/health"
