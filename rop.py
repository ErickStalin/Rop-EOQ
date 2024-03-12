import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from datetime import datetime
import math

class DataViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Datos")

        # Configurar el tamaño de la ventana principal
        self.geometry("800x600")

        self.frame = ttk.Frame(self)
        self.frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.table = ttk.Treeview(self.frame)
        self.table.pack(side="left", fill="both", expand=True)

        # Barra de desplazamiento horizontal
        self.scroll_x = tk.Scrollbar(self.frame, orient="horizontal", command=self.table.xview)
        self.scroll_x.pack(side="bottom", fill="x")
        self.table.configure(xscrollcommand=self.scroll_x.set)

        # Botón para cargar el archivo de Excel
        self.load_button = ttk.Button(self, text="Cargar archivo Excel", command=self.load_data)
        self.load_button.pack(pady=10)

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

            self.table["columns"] = list(datos.columns)
            for column in self.table["columns"]:
                self.table.column(column, anchor="center", width=100)
                self.table.heading(column, text=column)

            self.table.column("#0", width=50)  # Ancho de la columna del índice

            self.table.delete(*self.table.get_children())
            for index, row in datos.iterrows():
                self.table.insert("", "end", text=index, values=list(row))

if __name__ == "__main__":
    app = DataViewer()
    app.mainloop()
