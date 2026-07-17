# cruce-service

Microservicio de procesos pesados de Leoncito (antes corrían en el navegador).

## Endpoints
- `GET /health` — estado del servicio.
- `POST /limpieza/sedes` — dedup de teléfonos por DNI.
  - Body: **multipart** con campo `archivo` (Excel) **o** JSON `{ "filas": [ ... ] }`.
  - Respuesta JSON: `{ ok, stats, headers, filas, preview }`.
  - `?format=xlsx` → descarga directa del Excel limpio (`base_limpia.xlsx`).

## Correr en local (sin Docker)
```bash
npm install
npm start          # http://localhost:4002
```

## Correr con Docker
```bash
docker build -t cruce-service .
docker run -p 4002:4002 cruce-service
```

## Prueba rápida
```bash
curl http://localhost:4002/health
curl -F "archivo=@base.xlsx" http://localhost:4002/limpieza/sedes
```

## Roadmap
- [ ] Cruce cartera ↔ gestión por DNI + teléfono (avance-cartera, embudos, comparativo).
- [ ] Modo "call" de limpieza (quita NO CONTACTO cruzando con la gestión).
