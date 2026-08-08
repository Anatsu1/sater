#!/bin/bash
#------------------------------------------------------------------------------
# Compila y ejecuta SATER con Mono en Linux/macOS.
#
# No usa msbuild: el .vbproj arrastra el Application Framework de Visual Basic
# (My.MyApplication) y el compilador vbnc de Mono no lo implementa. En su lugar
# se invoca vbnc directamente con:
#
#   - herramientas/mono/Arranque.vb como punto de entrada (reemplaza a
#     My.MyApplication, que Visual Studio sigue usando sin cambios),
#   - los .resx recompilados sin el icono de ventana, porque libgdiplus no puede
#     convertir el .ico de Windows y la app aborta al crear la primera ventana,
#   - My Project/Settings.Designer.vb excluido: es boilerplate de Visual Studio
#     que la aplicacion no usa y que vbnc no sabe parsear.
#
# Requisitos: mono-complete, mono-basic (vbnc), python3 y, opcionalmente,
# mdbtools para migrar los datos reales del .mdb.
#
# Uso:  bash herramientas/mono/compilar.sh [--ejecutar]
#------------------------------------------------------------------------------
set -euo pipefail

RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
SRC="$RAIZ/PanelesSolares"
OUT="$SRC/bin/Debug"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

LIB="${MONO_LIB:-/usr/lib/mono/4.5}"
[ -d "$LIB" ] || { echo "No se encontro $LIB. Instala mono-complete o define MONO_LIB."; exit 1; }
[ -f "$LIB/vbnc.exe" ] || { echo "Falta vbnc.exe. Instala mono-basic."; exit 1; }

echo ">> Generando base SQLite desde el .mdb original"
python3 "$RAIZ/herramientas/migrar_mdb_a_sqlite.py" \
    "$OUT/paneleSolares.mdb" "$OUT/paneleSolares.db"

echo ">> Recompilando recursos sin el icono de ventana"
python3 - "$SRC" "$TMP" <<'PY'
import sys, glob, os
import xml.etree.ElementTree as ET
src, dest = sys.argv[1], sys.argv[2]
for f in glob.glob(os.path.join(src, '*.resx')):
    arbol = ET.parse(f)
    raiz = arbol.getroot()
    for nodo in list(raiz.findall('data')):
        if nodo.get('name', '').endswith('$this.Icon'):
            raiz.remove(nodo)
    arbol.write(os.path.join(dest, os.path.basename(f)),
                encoding='utf-8', xml_declaration=True)
PY

for f in "$TMP"/*.resx; do
    base="$(basename "$f" .resx)"
    mono "$LIB/resgen.exe" "$f" "$TMP/PanelesSolares.$base.resources" > /dev/null
done
mono "$LIB/resgen.exe" "$SRC/My Project/Resources.resx" \
     "$TMP/PanelesSolares.Resources.resources" > /dev/null

RES=""
for r in "$TMP"/*.resources; do RES="$RES /resource:$r"; done

echo ">> Compilando"
cd "$SRC"
# shellcheck disable=SC2086
mono "$LIB/vbnc.exe" \
  /target:winexe \
  /out:"$OUT/PanelesSolares.exe" \
  /rootnamespace:PanelesSolares \
  /main:PanelesSolares.Arranque \
  /define:'_MYTYPE=\"WindowsForms\"' \
  /optionstrict- /optionexplicit+ /optioninfer+ /optioncompare:binary \
  /reference:"$LIB/System.dll" \
  /reference:"$LIB/System.Data.dll" \
  /reference:"$LIB/System.Drawing.dll" \
  /reference:"$LIB/System.Windows.Forms.dll" \
  /reference:"$LIB/System.Xml.dll" \
  /reference:"$LIB/System.Core.dll" \
  /reference:"$LIB/System.Configuration.dll" \
  /reference:"$LIB/System.Deployment.dll" \
  /reference:"$LIB/Microsoft.VisualBasic.dll" \
  /imports:Microsoft.VisualBasic,System,System.Collections,System.Collections.Generic,System.Data,System.Drawing,System.Diagnostics,System.Windows.Forms,System.Linq,System.Xml.Linq \
  $RES \
  ./*.vb "My Project/AssemblyInfo.vb" "My Project/Resources.Designer.vb" \
  "$RAIZ/herramientas/mono/Arranque.vb"

echo ">> Listo: $OUT/PanelesSolares.exe"

if [ "${1:-}" = "--ejecutar" ]; then
    cd "$OUT" && exec mono PanelesSolares.exe
fi
