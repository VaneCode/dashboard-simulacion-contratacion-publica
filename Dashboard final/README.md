# Dashboard PAC–POA 2026

Dashboard académico creado con HTML, CSS y JavaScript DOM. Consume los mismos archivos JSON generados por el motor de simulación de la tesis.

## Ejecución local

```bash
python -m http.server 8080 --directory dist
```

Abrir `http://localhost:8080`.

## Datos

- `pac_limpio.json`: fuente PAC observada y campos derivados.
- `ciclo_sintetico.json`: variables sintéticas de contratación.
- `sensibilidad_resumen.json`: resumen de 100 repeticiones por escenario.

Los datos sintéticos no representan el desempeño real del GAD Municipal de Guayaquil.
