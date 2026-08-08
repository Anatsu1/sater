#!/usr/bin/env python3
"""
Migra la base Access original (paneleSolares.mdb) a SQLite (paneleSolares.db).

El sistema fue escrito contra Microsoft Jet OLEDB 4.0, que solo existe en
Windows. Para poder ejecutar y demostrar el proyecto en Linux/macOS, esta
herramienta vuelca las dos tablas del .mdb a un archivo SQLite equivalente, que
es el que usa la aplicacion cuando BaseDatos detecta que no corre sobre Windows.

Uso:
    python3 herramientas/migrar_mdb_a_sqlite.py \
        PanelesSolares/bin/Debug/paneleSolares.mdb \
        PanelesSolares/bin/Debug/paneleSolares.db

Requiere mdbtools (mdb-export). Si no esta instalado, cae en un juego de datos
de referencia tomado de las planillas de tablas/.
"""

import os
import shutil
import sqlite3
import subprocess
import sys
import csv
import io

ESQUEMA = """
DROP TABLE IF EXISTS electro;
CREATE TABLE electro (
    idElectro INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre    TEXT NOT NULL,
    hsConsumo REAL NOT NULL,
    conKwh    REAL NOT NULL
);

DROP TABLE IF EXISTS paneles;
CREATE TABLE paneles (
    idPanel     INTEGER PRIMARY KEY AUTOINCREMENT,
    tipoPanel   TEXT NOT NULL,
    watts       REAL NOT NULL,
    eficiencia  REAL NOT NULL,
    marca       TEXT NOT NULL,
    modelo      TEXT NOT NULL,
    dimensiones TEXT NOT NULL
);
"""

# Respaldo usado cuando mdbtools no esta disponible. Sale de tablas/panelSolar.xlsx
# y tablas/tablapaneles.xlsx, que son las planillas de relevamiento del proyecto.
ELECTRO_RESPALDO = [
    ("Heladera", 24, 0.06),
    ("Heladera con freezer", 24, 0.09),
    ("Freezer", 24, 0.085),
    ("Microondas", 2, 0.64),
    ("Lavarropas automatico de 5 kg", 3, 0.17),
    ("Secarropas centrifugo", 1, 0.19),
    ("Secarropas a calor", 1, 0.35),
    ("Luces", 6, 0.3),
    ("Computadora (solo la CPU)", 5, 0.72),
    ("Televisor LED 24\"", 3, 0.06),
    ("Televisor LED 32\" a 50\"", 3, 0.1),
    ("Televisor LCD de 40\"", 3, 0.15),
    ("Reproductor Video", 8, 0.6),
    ("Aire Acondicionado", 3, 1.013),
    ("Radiador", 6, 0.96),
    ("Ventilador de pie", 3, 0.7),
    ("Ventilador de techo", 3, 0.075),
    ("Estufa de cuarzo c/termostato", 4, 1.2),
    ("Termotanque", 5, 0.9),
    ("Horno electrico", 1, 1.5),
]

PANELES_RESPALDO = [
    ("amorfo", 100, 20, "Sunpower", "Sunpower Mini", "10x10x20"),
    ("amorfo", 150, 21, "Sunpower", "Sunpower Mini", "20x25x30"),
    ("policristalino", 200, 22, "Sunpower", "Sunpower pro", "20x25x35"),
    ("policristalino", 250, 15.1, "Solax", "panel solar solax policristalino", "1332x992x35"),
    ("monocristalino", 300, 24, "Sunpower", "Sunpower pro", "20x25x40"),
    ("monocristalino", 460, 20.7, "JA Solar", "Panel JA Solar Monocristalino Perc", "2120x1052x35"),
    ("Placa solar bifacial", 514, 30, "LG", "LG NeON 2 BiFacial", "2064x1024x40"),
    ("Panel solar de celula PERC", 670, 30, "sunev", "Panel Solar Fotovoltaico PERC", "2384x1303x35"),
    ("monocristalino", 1000, 26, "Sunpower", "Sunpower pro", "300x300x100"),
    ("monocristalino", 1500, 27, "Sunpower", "Sunpower pro", "400x400x100"),
    ("Placa solar de pelicula delgada", 3000, 50, "SolarEdge", "Inversor SolarEdge HD-Wave SE3000H", "142x370x280"),
]


def exportar_con_mdbtools(mdb, tabla):
    """Devuelve las filas de `tabla` como lista de dicts, o None si falla."""
    if shutil.which("mdb-export") is None:
        return None
    try:
        salida = subprocess.run(
            ["mdb-export", mdb, tabla],
            capture_output=True, text=True, check=True,
        ).stdout
    except (subprocess.CalledProcessError, UnicodeDecodeError):
        return None
    filas = list(csv.DictReader(io.StringIO(salida)))
    return filas or None


def numero(valor, por_defecto=0.0):
    try:
        return float(str(valor).replace(",", "."))
    except (TypeError, ValueError):
        return por_defecto


def main():
    mdb = sys.argv[1] if len(sys.argv) > 1 else "PanelesSolares/bin/Debug/paneleSolares.mdb"
    destino = sys.argv[2] if len(sys.argv) > 2 else "PanelesSolares/bin/Debug/paneleSolares.db"

    if os.path.exists(destino):
        os.remove(destino)

    cn = sqlite3.connect(destino)
    cn.executescript(ESQUEMA)

    electro = exportar_con_mdbtools(mdb, "electro") if os.path.exists(mdb) else None
    if electro:
        datos = [
            (f.get("nombre", "").strip(),
             numero(f.get("hsConsumo")),
             numero(f.get("conKwh")))
            for f in electro if f.get("nombre", "").strip()
        ]
        origen_electro = "mdbtools"
    else:
        datos = ELECTRO_RESPALDO
        origen_electro = "respaldo"
    cn.executemany("INSERT INTO electro (nombre, hsConsumo, conKwh) VALUES (?, ?, ?)", datos)

    paneles = exportar_con_mdbtools(mdb, "paneles") if os.path.exists(mdb) else None
    if paneles:
        datos_p = [
            (f.get("tipoPanel", "").strip(),
             numero(f.get("watts")),
             numero(f.get("eficiencia")),
             f.get("marca", "").strip(),
             f.get("modelo", "").strip(),
             f.get("dimensiones", "").strip())
            for f in paneles if f.get("tipoPanel", "").strip()
        ]
        origen_paneles = "mdbtools"
    else:
        datos_p = PANELES_RESPALDO
        origen_paneles = "respaldo"
    cn.executemany(
        "INSERT INTO paneles (tipoPanel, watts, eficiencia, marca, modelo, dimensiones) "
        "VALUES (?, ?, ?, ?, ?, ?)", datos_p)

    cn.commit()
    e = cn.execute("SELECT COUNT(*) FROM electro").fetchone()[0]
    p = cn.execute("SELECT COUNT(*) FROM paneles").fetchone()[0]
    cn.close()

    print("Base generada: %s" % destino)
    print("  electro: %d filas (origen: %s)" % (e, origen_electro))
    print("  paneles: %d filas (origen: %s)" % (p, origen_paneles))


if __name__ == "__main__":
    main()
