import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
import math
import sqlite3

class DataViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Datos")

        # Configurar el tamaño de la ventana principal
        self.geometry("800x600")

        self.frame = ttk.Frame(self)
        self.frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.search_frame = ttk.Frame(self.frame)
        self.search_frame.pack(pady=10, fill="x")

        self.search_label = ttk.Label(self.search_frame, text="Buscar por Nombre:")
        self.search_label.pack(side="left")

        self.search_entry = ttk.Entry(self.search_frame)
        self.search_entry.pack(side="left", padx=(5, 0), fill="x", expand=True)

        self.search_button = ttk.Button(self.search_frame, text="Buscar", command=self.search_data)
        self.search_button.pack(side="left", padx=(5, 0))

        self.table = ttk.Treeview(self.frame)
        self.table.pack(side="left", fill="both", expand=True)

        # Barra de desplazamiento horizontal
        self.scroll_x = tk.Scrollbar(self.frame, orient="horizontal", command=self.table.xview)
        self.scroll_x.pack(side="bottom", fill="x")
        self.table.configure(xscrollcommand=self.scroll_x.set)

        # Botón para cargar el archivo de Excel
        self.load_button = ttk.Button(self, text="Cargar archivo Excel", command=self.load_data)
        self.load_button.pack(side="left", padx=10, pady=10)

        # Botón para visualizar datos desde la base de datos
        self.visualize_button = ttk.Button(self, text="Visualizar data", command=self.visualize_data)
        self.visualize_button.pack(side="left", padx=10, pady=10)

        # Campo de entrada para notas
        self.note_entry = ttk.Entry(self.frame)
        self.note_entry.pack(side="top", pady=(10, 0), fill="x", expand=True)

        # Botón para agregar nota
        self.add_note_button = ttk.Button(self.frame, text="Agregar Nota", command=self.add_note)
        self.add_note_button.pack(side="top", pady=(5, 0))

        # Cargar datos desde la base de datos al iniciar la aplicación
        self.load_data_from_db()

    def load_data(self):
        # Diálogo para seleccionar el archivo de Excel
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        if file_path:
            datos = pd.read_excel(file_path)

            def convertir_nat_a_vacio(valor):
                return '' if pd.isnull(valor) or pd.isna(valor) else valor

            columnas_fechas = ['FechaIngreso', 'FechaÚltimoIngreso']
            datos[columnas_fechas] = datos[columnas_fechas].applymap(convertir_nat_a_vacio)

            datos['Ingresos'] = datos['Stock'] + datos['Vendido']

            def calcular_rotacion_mensual(vendido, fecha_ingreso):
                if vendido > 0:
                    dias_transcurridos = (datetime.now() - fecha_ingreso).days
                    if dias_transcurridos > 0:
                        return round((vendido / dias_transcurridos) * 30, 2)
                return ""

            datos['RotaciónMensual'] = datos.apply(lambda row: calcular_rotacion_mensual(row['Vendido'], row['FechaIngreso']), axis=1)

            def calcular_rotura_stock(tiempo_entrega_dias, vendido, fecha_ingreso):
                if vendido > 0:
                    dias_transcurridos = (datetime.now() - fecha_ingreso).days
                    if dias_transcurridos > 0:
                        return round(tiempo_entrega_dias * (vendido / dias_transcurridos), 0)
                return ""

            datos['RoturaStock'] = datos.apply(lambda row: calcular_rotura_stock(row['TiempoEntregaDías'], row['Vendido'], row['FechaIngreso']), axis=1)

            def calcular_estrategia_compra(rotura_stock, stock, tiempo_entrega_dias, vendido, fecha_ingreso):
                if rotura_stock == "":
                    return ""
                elif rotura_stock > 0:
                    dias_transcurridos = (datetime.now() - fecha_ingreso).days
                    if dias_transcurridos > 0:
                        nivel_reposicion = stock / (tiempo_entrega_dias * (vendido / dias_transcurridos))
                        if nivel_reposicion <= 1:
                            return "Reordenar"
                        elif nivel_reposicion <= 1.25:
                            return "Preparar"
                return ""

            datos['EstrategiaCompra'] = datos.apply(lambda row: calcular_estrategia_compra(row['RoturaStock'], row['Stock'], row['TiempoEntregaDías'], row['Vendido'], row['FechaIngreso']), axis=1)

            def calcular_costo_mantener(costo):
                return costo * 26 / 100

            datos['CostoMantener'] = datos['Costo'].apply(calcular_costo_mantener)

            def calcular_cantidad_reorden(estrategia_compra, costo_mantener, rotura_stock, existencias_totales, vendido, costo_ordenar):
                if estrategia_compra in ["Reordenar", "Preparar"]:
                    if costo_mantener == "":
                        return ""
                    elif costo_mantener > 0:
                        return round((rotura_stock - existencias_totales) + math.sqrt((2 * vendido * costo_ordenar) / costo_mantener), 0)
                return ""

            datos['CantidadReorden'] = datos.apply(lambda row: calcular_cantidad_reorden(row['EstrategiaCompra'], row['CostoMantener'], row['RoturaStock'], row['Stock'], row['Vendido'], row['CostoOrdenar']), axis=1)

            # Agregar la columna "Codigo" después de la columna "NombreP"
            datos.insert(1, "Codigo", range(1, len(datos) + 1))

            self.original_data = datos.copy()

            self.table["columns"] = list(datos.columns)
            for column in self.table["columns"]:
                self.table.column(column, anchor="center", width=100)
                self.table.heading(column, text=column)

            self.table.column("#0", width=50)  # Ancho de la columna del índice

            self.table.delete(*self.table.get_children())
            for index, row in datos.iterrows():
                self.table.insert("", "end", text=index, values=list(row))

            # Guardar los datos en la base de datos
            self.guardar_en_base_de_datos(datos)

    def guardar_en_base_de_datos(self, datos):
        conn = sqlite3.connect("datos_calculados.db")
        c = conn.cursor()

        c.execute('''CREATE TABLE IF NOT EXISTS Tabla1 (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        NombreP TEXT,
                        Stock REAL,
                        Vendido REAL,
                        Ingresos REAL,
                        RotaciónMensual REAL,
                        RoturaStock REAL,
                        EstrategiaCompra TEXT,
                        CostoMantener REAL,
                        CantidadReorden REAL,
                        Notas TEXT
                    )''')

        for index, row in datos.iterrows():
            c.execute('''INSERT INTO Tabla1 (NombreP, Stock, Vendido, Ingresos, RotaciónMensual, RoturaStock, EstrategiaCompra, CostoMantener, CantidadReorden, Notas)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', (row['NombreP'], row['Stock'], row['Vendido'], row['Ingresos'], row['RotaciónMensual'], row['RoturaStock'], row['EstrategiaCompra'], row['CostoMantener'], row['CantidadReorden'], ""))

        conn.commit()
        conn.close()

    def visualize_data(self):
        try:
            conn = sqlite3.connect("datos_calculados.db")
            c = conn.cursor()

            c.execute("SELECT * FROM Tabla1")
            rows = c.fetchall()

            self.table.delete(*self.table.get_children())
            for row in rows:
                self.table.insert("", "end", values=row)

            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo visualizar los datos: {e}")

    def search_data(self):
        query = self.search_entry.get().strip()
        if query:
            results = self.original_data[self.original_data['NombreP'].str.contains(query, case=False)]
            self.table.delete(*self.table.get_children())
            for index, row in results.iterrows():
                self.table.insert("", "end", text=index, values=list(row))
        else:
            self.table.delete(*self.table.get_children())
            for index, row in self.original_data.iterrows():
                self.table.insert("", "end", text=index, values=list(row))

    def load_data_from_db(self):
        try:
            conn = sqlite3.connect("datos_calculados.db")
            datos = pd.read_sql_query("SELECT * FROM Tabla1", conn)

            self.original_data = datos.copy()

            self.table["columns"] = list(datos.columns)
            for column in self.table["columns"]:
                self.table.column(column, anchor="center", width=100)
                self.table.heading(column, text=column)

            self.table.column("#0", width=50)  # Ancho de la columna del índice

            self.table.delete(*self.table.get_children())
            for index, row in datos.iterrows():
                self.table.insert("", "end", text=index, values=list(row))

            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar datos desde la base de datos: {e}")

    def add_note(self):
        selected_item = self.table.focus()
        if selected_item:
            note = self.note_entry.get().strip()
            if note:
                note_index = int(selected_item.lstrip("I"))
                current_notes = self.table.item(selected_item, "values")[-1]
                new_notes = current_notes + "\n" + note if current_notes else note
                self.table.item(selected_item, values=(self.table.item(selected_item, "values")[:-1] + (new_notes,)))
                conn = sqlite3.connect("datos_calculados.db")
                c = conn.cursor()
                c.execute("UPDATE Tabla1 SET Notas = ? WHERE Id = ?", (new_notes, note_index))
                conn.commit()
                conn.close()
                self.note_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("Advertencia", "Por favor ingresa una nota antes de agregarla.")
        else:
            messagebox.showwarning("Advertencia", "Por favor selecciona un elemento de la tabla antes de agregar una nota.")

if __name__ == "__main__":
    app = DataViewer()
    app.mainloop()
