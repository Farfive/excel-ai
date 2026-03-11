#!/bin/bash
set -e

docker-compose down -v
rm -rf chroma_db/* models/*
echo "Reset complete. Run ./scripts/setup.sh"
