#!/usr/bin/env python3
"""Descarga los DOS archivos oficiales a assets/ y reconstruye QMD/HTML con rutas locales.

Requiere conexión a Internet solamente en el momento de la descarga.
Tras descargar, la presentación usa las imágenes locales al volver a generar QMD/HTML.

Uso desde la raíz del proyecto: python tools/cachear_logos.py
"""
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError
import subprocess, sys

ROOT=Path(__file__).resolve().parent.parent
LOGOS=[
 ('https://congresociie.cl/wp-content/uploads/2026/04/CIIE-logo-septimo.png','assets/logo-ciie-septimo.png',b'\x89PNG\r\n\x1a\n'),
 ('https://facso.uchile.cl/dam/jcr:16208a63-f7b7-486f-9c82-85eed075421e/logoicono150.jpg','assets/logo-facso-uchile.jpg',b'\xff\xd8\xff'),
]

for url,rel,magic in LOGOS:
 dest=ROOT/rel
 if dest.exists() and dest.read_bytes().startswith(magic):
  print('Ya existe:', rel)
  continue
 try:
  req=Request(url,headers={'User-Agent':'Mozilla/5.0 (compatible; CIIE-2026-presentation/1.0)'})
  with urlopen(req,timeout=20) as response:
   data=response.read(5_000_000)
  if not data.startswith(magic):
   print(f'Error: la respuesta no es una imagen válida ({rel})',file=sys.stderr)
   continue
  dest.parent.mkdir(parents=True,exist_ok=True)
  dest.write_bytes(data)
  print('Descargado:', rel, len(data), 'bytes')
 except (OSError,URLError,ValueError) as exc:
  print('No fue posible descargar',url,'\n',exc,file=sys.stderr)

if all((ROOT/rel).exists() and (ROOT/rel).read_bytes().startswith(magic) for _,rel,magic in LOGOS):
 print('Ambos logotipos oficiales están disponibles localmente.')
 print('Actualizando automáticamente ponencia.qmd y ponencia.html ...')
 subprocess.run([sys.executable,str(ROOT/'build_from_v29.py')],cwd=ROOT,check=True)
else:
 print('Descarga incompleta: las imágenes aún referencian los servidores oficiales.')
 sys.exit(2)
