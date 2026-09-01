# COFEPRIS — Registros sanitarios de medicamentos

Interfaz web (Streamlit) sobre la lógica del notebook `COFEPRIS_final_1.ipynb`:
descarga la base oficial de otorgados del visor de COFEPRIS, aplica filtros
configurables (ATC, texto, estado, tipo, titular, fabricante, fechas), muestra
la serie mensual y permite descargar los resultados en Excel y CSV.

La parte de solicitudes (PDFs de gob.mx) no está incluida: la mayoría de los
PDFs están cifrados y gob.mx bloquea el scraping. Para esa serie, tramita el
dato por Plataforma Nacional de Transparencia.

## Correr en tu máquina

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en http://localhost:8501.

## Desplegar en Streamlit Community Cloud (gratis)

1. Crea un repositorio en GitHub (puede ser privado) y sube estos 3 archivos:
   `app.py`, `requirements.txt` y `README.md`.
2. Entra a https://share.streamlit.io e inicia sesión con tu cuenta de GitHub.
3. Clic en **Create app** → **Deploy a public app from GitHub**.
4. Elige el repositorio, rama `main` y archivo principal `app.py`.
5. Clic en **Deploy**. En un par de minutos tendrás tu URL pública
   (algo como `https://tu-app.streamlit.app`).

Cada `git push` al repositorio redespliega la app automáticamente.

## Logo y tema de Pfizer

- La app busca el logo en `assets/logo_pfizer.png` (o `.svg`) dentro del repo;
  si no lo encuentra, usa el logo oficial publicado en Wikimedia Commons.
  Lo recomendable es subir al repo el archivo oficial de brand assets de Pfizer.
- Los colores corporativos (azul #0000C9, azul claro #3D96F7) pueden configurarse en
  `.streamlit/config.toml`.

## Notas

- La base se descarga del visor oficial y se cachea 6 horas; el botón de menú
  de Streamlit (⋮ → Rerun) o esperar el vencimiento del caché la refresca.
- Si la descarga falla (COFEPRIS caído o bloqueado), la app ofrece subir el
  XLSX manualmente.
- El botón **Preset oncología** aplica el mismo filtro del notebook:
  ATC L01 + L02, términos de cáncer y rescate de registros sin clasificación
  ATC (sin esto se pierden Tamoxifeno, Enzalutamida y ~109 más).
