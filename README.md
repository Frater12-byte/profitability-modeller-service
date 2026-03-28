# Profitability Modeller Service
Flask microservice that rebuilds the Profitability_Modeller_2026.xlsx from PowerBI exports.

## Endpoints
- GET /health — health check
- POST /rebuild — accepts {data1_b64, data2_b64}, returns {modeller_b64}

