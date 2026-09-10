# Dashboard y simulación PAC–POA 2026 — Contratación pública GAD Municipal de Guayaquil

Repositorio de titularidad exclusiva de la autora que reúne el **código de simulación**
y el **dashboard de visualización** utilizados en la tesis *"Modelo integrado
PAC–POA y ciclo de contratación pública 2026"*.

> **Advertencia metodológica:** el PAC constituye la fuente observada. Los
> objetivos POA, las fechas, los estados y los montos de adjudicación y
> ejecución son **sintéticos**, generados exclusivamente con fines académicos.
> No representan el desempeño real del GAD Municipal de Guayaquil.

## Estructura del repositorio

```
.
├── Codigo para simulacion/         # Motor de simulación (Python)
│   ├── generate_tesis_data.py      # Depura el PAC, genera POA sintético y simula el ciclo contractual
│   ├── generar_modelo_excel.py     # Exporta los resultados a un libro Excel auditable
│   ├── requirements_tesis.txt      # Dependencias (xlrd, openpyxl)
│   ├── upload/                     # PAC original (fuente observada)
│   ├── generated_data/             # Salidas JSON del modelo (PAC limpio, POA y ciclo sintéticos, sensibilidad)
│   └── Modelo_PAC_POA_Contratacion_2026_reproducido.xlsx
│
├── Dashboard final/                # Visualización (HTML + CSS + JS, sin frameworks)
│   └── dist/
│       ├── index.html, app.js, styles.css, charts.css
│       └── data/                   # Copia de los JSON que consume el dashboard
│
├── .github/workflows/deploy.yml    # Publica automáticamente "Dashboard final/dist" en GitHub Pages
├── LICENSE
└── README.md
```

## Reproducibilidad

- **Semilla:** `2026`
- **Escenarios:** Conservador, Base, Favorable
- **Repeticiones del análisis de sensibilidad:** 100 por escenario
- **Registros esperados del PAC depurado:** 736
- **Resultado base esperado:** materialización 79,89 %, finalización 74,32 %,
  mediana de 38 días hasta adjudicación

## Cómo reproducir la simulación

```bash
cd "Codigo para simulacion"
python -m venv .venv
source .venv/bin/activate        # En Windows: .venv\Scripts\activate
pip install -r requirements_tesis.txt
python generate_tesis_data.py
python generar_modelo_excel.py
```

## Cómo ver el dashboard

**Local:**
```bash
cd "Dashboard final/dist"
python -m http.server 8080
# abrir http://localhost:8080
```

**En línea (GitHub Pages):** el workflow incluido en
`.github/workflows/deploy.yml` publica automáticamente la carpeta
`Dashboard final/dist` cada vez que se hace push a `main`. Una vez activado
(ver sección "Despliegue" más abajo), el dashboard queda disponible en:

```
https://<usuario>.github.io/<nombre-del-repositorio>/
```

## Datos

Ver `Codigo para simulacion/README_CODIGO_TESIS.md` y `Dashboard final/README.md`
para el detalle de cada archivo de datos. Los datos sintéticos están rotulados
como tales y son completamente regenerables con la semilla `2026`; no contienen
datos personales.

## Licencia

Código bajo licencia MIT (ver `LICENSE`). Los datos derivados/sintéticos se
distribuyen únicamente con fines académicos, sin garantía de exactitud
respecto a la ejecución real de la entidad.
