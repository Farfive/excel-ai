#!/bin/bash
set -e

echo "=== Health Check ==="
curl -s http://localhost:8000/health | python3 -m json.tool

echo ""
echo "=== Upload Test XLSX ==="
RESPONSE=$(curl -s -X POST http://localhost:8000/workbook/upload \
  -F "file=@tests/fixtures/sample.xlsx" \
  -H "Accept: application/json")
echo "$RESPONSE" | python3 -m json.tool
UUID=$(echo "$RESPONSE" | python3 -c "import sys,json; print(json.load(sys.stdin)['workbook_uuid'])")

echo ""
echo "=== Ask Question (SSE) ==="
curl -s -X POST "http://localhost:8000/workbook/${UUID}/ask" \
  -H "Content-Type: application/json" \
  -d '{"question": "What are the key assumptions in this model?"}' \
  --no-buffer | head -20

echo ""
echo "=== Anomalies ==="
curl -s "http://localhost:8000/workbook/${UUID}/anomalies" | python3 -m json.tool
