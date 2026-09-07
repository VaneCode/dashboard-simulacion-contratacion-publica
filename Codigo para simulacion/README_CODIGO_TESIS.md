# Código de simulación y generación de datos de la tesis

Este paquete reproduce el modelo PAC–POA–ciclo contractual utilizado en la tesis.

## Contenido

- `generate_tesis_data.py`: depura el PAC, genera el POA sintético, simula el ciclo contractual y ejecuta 100 repeticiones para cada escenario.
- `generar_modelo_excel.py`: transforma las salidas JSON en un libro Excel auditable.
- `requirements_tesis.txt`: dependencias de Python.
- `upload/PAC 2026 Guayaquil.xls`: fuente observada utilizada como entrada.
- `generated_data/`: salidas generadas por el modelo.

## Requisitos

- Python 3.10 o superior.
- `pip` habilitado.

## Instalación y ejecución

Desde la carpeta descomprimida, ejecutar:

```bash
python -m venv .venv
```

En Windows:

```bash
.venv\Scripts\activate
pip install -r requirements_tesis.txt
python generate_tesis_data.py
python generar_modelo_excel.py
```

En Linux o macOS:

```bash
source .venv/bin/activate
pip install -r requirements_tesis.txt
python generate_tesis_data.py
python generar_modelo_excel.py
```

## Reproducibilidad

- Semilla de la corrida base: `2026`.
- Escenarios: Conservador, Base y Favorable.
- Repeticiones del análisis de sensibilidad: 100 por escenario.
- Registros esperados del PAC depurado: 736.
- Resultado base esperado: materialización 79,89 %, finalización 74,32 % y mediana de 38 días hasta adjudicación.

## Advertencia metodológica

El PAC constituye la fuente observada. Los objetivos POA, las fechas, los estados y los montos de adjudicación y ejecución son sintéticos. Se generan exclusivamente para fines académicos y no representan la ejecución real ni el desempeño del GAD Municipal de Guayaquil.
