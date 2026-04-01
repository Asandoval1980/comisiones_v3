# Liquidador de comisiones v3

## Qué cambia en esta versión
- Usa reglas fijas internas para la liquidación.
- Ya no pide el archivo de tabla de comisiones.
- Exporta el resultado final con la estructura base del archivo ejemplo `Liquidación de Comisiones`.
- Mantiene solo las hojas originales del formato final.
- Barranquilla sigue fuera del cálculo automático.

## Ejecución
```bash
pip install -r requirements.txt
streamlit run app.py
```
