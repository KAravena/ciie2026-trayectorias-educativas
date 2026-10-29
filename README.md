# P47.4 — Quarto Reveal.js CIIE 2026 · 19 diapositivas

## Qué cambió respecto de P47.3

Se **eliminó completamente** de la presentación la diapositiva titulada
«Cambiar las categorías puede cambiar la lectura de IP», antes numerada 19.

- Las **diapositivas 1–18** permanecen sin cambios científicos ni visuales.
- **Referencias** pasa a ser la **diapositiva 19** (última).
- `ponencia.qmd` contiene 19 encabezados Reveal.js y mantiene las notas del expositor.
- `ponencia.html` contiene 19 secciones, los controles de teclado/ratón y el contador dinámico `1 / 19`.
- `ponencia.pdf` contiene 19 páginas, con la lámina de sensibilidad de IP completamente ausente.
- `referencias.html` vuelve a la última diapositiva (`ponencia.html#/18`).
- `build_from_v29.py` **omite la diapositiva original 19** al reconstruir el proyecto; volver a ejecutar ese script no restaura la diapositiva retirada.
- `capturas/` ha sido regenerada para mostrar solo las 19 páginas vigentes.

## Archivos para utilizar

- `ponencia.qmd`: archivo fuente editable para Quarto Reveal.js.
- `ponencia.html`: reproductor local autónomo, estilos y JavaScript incluidos; abre directamente en Chromium/Chrome/Edge, sin Quarto. No es el runtime generado por `quarto render`.
- `ponencia.pdf`: respaldo listo para proyectar, 19 páginas.
- `referencias.html`, `referencias.pdf`: bibliografía completa; enlace desde la última diapositiva.
- `styles/`: tema SCSS y CSS para renderizado nativo por Quarto.
- `assets/`: runtime local y datos auxiliares.
- `data/`: resultados de referencia.
- `fuentes/`: material histórico V29 del cual se reconstruyó el proyecto. **El PPTX original se conserva sin modificar y no debe confundirse con la presentación final de 19 diapositivas**.
- `capturas/`: capturas actualizadas de las 19 páginas y hoja de contacto.

## Presentar y compilar

Para presentar sin instalar nada, abrir `ponencia.html`. Flechas o espacio para avanzar; `S` para notas; `F` para pantalla completa. La ponencia termina en la diapositiva 18 y la 19 muestra referencias.

Para compilar el QMD nativo, desde este directorio:

```bash
quarto render ponencia.qmd --to revealjs
```

La salida se configura en `_site/`. En este entorno no estaba disponible el CLI de Quarto, por lo que **el HTML local está verificado, pero el renderizado nativo en Quarto queda pendiente**. Es necesario probar la salida nativa antes de presentar con ella.

**Logotipos:** portada conserva enlaces oficiales, con texto de respaldo. Se requiere conexión o ejecutar `python tools/cachear_logos.py` para guardar imágenes oficiales localmente. El PDF de respaldo conserva las páginas de la versión P47.3 con la página 19 eliminada, sin alterar otros elementos.

## Integridad

La exclusión de una lámina no modifica los análisis ni sus resultados. El QMD mantiene la secuencia M1–M3, estimaciones PPP, resultados multinomiales, cohortes y conclusiones; las notas de la lámina 18 permanecen como en la versión original. La sensibilidad IP sigue documentada en los materiales científicos originales, pero **ya no forma parte de la presentación**.
