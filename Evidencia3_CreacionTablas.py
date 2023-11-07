import datetime
import sqlite3
from sqlite3 import Error
try:
    with sqlite3.connect("Evidencia3.db") as conn:
        mi_cursor = conn.cursor()
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Clientes (ClaveC INTEGER PRIMARY KEY, NombreCompleto TEXT NOT NULL, RFC TEXT NOT NULL, CORREO TEXT NOT NULL);")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Servicios (ClaveS INTEGER PRIMARY KEY, Nombre TEXT NOT NULL, Costo REAL CHECK(Costo > 0));")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Notas (Folio INTEGER PRIMARY KEY, ClaveC INTEGER NOT NULL, Fecha TIMESTAMP, Estado_Nota INTEGER DEFAULT 0 ,FOREIGN KEY (ClaveC) REFERENCES Clientes(ClaveC));")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS DetalleNotas (Folio INTEGER, ClaveS INTEGER NOT NULL, Monto REAL, FOREIGN KEY (Folio) REFERENCES Notas(Folio),FOREIGN KEY (ClaveS) REFERENCES Servicios(ClaveS));")
        print("Tablas creadas")
except Error as e:
    print(e)
except:
    print("se produjo un error")