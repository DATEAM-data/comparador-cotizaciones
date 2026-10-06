# Comparador de Cotizaciones — Olivas & Asociados

App web: subes los PDFs de las cotizaciones (HDI, CHUBB, QUÁLITAS, GNP, INBURSA, ANA)
y descargas el COMPARATIVO.pdf (coberturas + hoja de opciones de pago).

## Archivos
- `app.py` — la página web (reemplaza a la Celda 3 de Colab)
- `comparador.py` — las funciones (tu Celda 2, con correcciones)
- `logos/` — solo necesitas `olivas.png`. Los logos de las aseguradoras se sacan solos de cada cotización.
  Si alguno sale mal o falta, pon un `chubb.png`, `gnp.png`, `qualitas.png`, `hdi.png`, `ana.png` o `inbursa.png` aquí y se usa como respaldo.
- `requirements.txt` — dependencias

## Publicarla gratis (Streamlit Community Cloud)
1. Crea un repositorio en GitHub y sube estos archivos (puede ser privado).
2. Entra a https://share.streamlit.io con tu cuenta de GitHub → **Create app**.
3. Elige el repositorio, rama `main`, archivo principal `app.py` → **Deploy**.
4. (Recomendado) En la app: **Settings → Secrets**, agrega:
   ```
   APP_PASSWORD = "tu-contraseña"
   ```
   Así solo entra quien tenga la contraseña.

Cada vez que cambies `comparador.py` en GitHub, la app se actualiza sola.

## Probar en tu computadora
```
pip install -r requirements.txt
streamlit run app.py
```
