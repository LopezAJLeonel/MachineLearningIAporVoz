# Proyecto: Inventario con Inteligencia Artificial
# Responsable: Jose Leonel Lopez Almeida
# Version: 3.0 Actualizada

# Fecha_Creacion: 06/06/2024
# Fecha_Modificacion: 10/08/2024
# Descripcion: 
# Este código implementa una aplicación de predicción de ventas utilizando una interfaz gráfica desarrollada en Tkinter. 
# La aplicación permite cargar archivos Excel, entrenar un modelo de predicción de ventas, visualizar gráficos de ventas 
# y predicciones, generar reportes en PDF, y realizar búsquedas mediante reconocimiento de voz. 
# También incluye funcionalidades para alertar al usuario sobre productos con stock bajo y visualizar las ventas por mes y tipo de producto.

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
import pandas as pd
import speech_recognition as sr
from modelo import SalesPredictor
from fpdf import FPDF
from fuzzywuzzy import process

class SalesPredictionApp:
    def __init__(self, root):
        # Inicializar la aplicacion
        self.root = root
        self.root.title("Predicción de Ventas")
        self.root.geometry("800x700")
        self.root.configure(bg='#1da58f')

        # Cargar y mostrar el logo de la aplicación
        self.logo = Image.open("imagen.jpg")
        self.logo = ImageTk.PhotoImage(self.logo)
        self.logo_label = ttk.Label(root, image=self.logo, background='white')
        self.logo_label.pack(pady=10)


        # Configurar estilos para los widgets
        style = ttk.Style()
        style.configure('TButton', font=('Arial', 12), padding=10, relief="flat", background='#1da58f', foreground='black')
        style.configure('TLabel', font=('Arial', 14), background='#1da58f', foreground='black')
        style.configure('TFrame', background='#1da58f')
        style.configure('Treeview', font=('Arial', 12), background='white', foreground='black', rowheight=25)
        style.configure('Treeview.Heading', font=('Arial', 12, 'bold'), background='#1da58f', foreground='black')

        # Crear un marco principal para contener los widgets
        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(expand=True, fill='both')

        # Botón para predecir ventas, inicialmente deshabilitado
        self.load_button = ttk.Button(main_frame, text="Cargar archivo Excel", command=self.load_file)
        self.load_button.pack(pady=10)

        self.predict_button = ttk.Button(main_frame, text="Predecir Ventas", command=self.predict_sales, state=tk.DISABLED)
        self.predict_button.pack(pady=10)

        # Imagen que actúa como un botón para la busqueda del voz
        self.voice_search_image = Image.open("ic.jpg")  # Asegúrate de tener esta imagen
        self.voice_search_image = self.voice_search_image.resize((50, 50), Image.LANCZOS)  # Redimensiona la imagen
        self.voice_search_image = ImageTk.PhotoImage(self.voice_search_image)

        # Botón de búsqueda por voz
        self.voice_search_button = tk.Button(main_frame, image=self.voice_search_image, command=self.voice_search, relief='flat', bg='white')
        self.voice_search_button.pack(pady=10)

        # Etiqueta para mostrar los resultados de la búsqueda por voz
        self.result_label = ttk.Label(main_frame, text="")
        self.result_label.pack(pady=10)

        # Configuración del Arbol para mostrar datos
        self.tree = ttk.Treeview(main_frame, columns=("Mes", "Tipo de Producto", "Unidades"), show='headings')
        self.tree.heading("Mes", text="Mes")
        self.tree.heading("Tipo de Producto", text="Tipo de Producto")
        self.tree.heading("Unidades", text="Unidades")
        self.tree.column("Mes", anchor=tk.CENTER)
        self.tree.column("Tipo de Producto", anchor=tk.CENTER)
        self.tree.column("Unidades", anchor=tk.CENTER)
        self.tree.pack(pady=20, fill='both', expand=True)

    def load_file(self):
        # Abre un cuadro de diálogo para seleccionar un archivo Excel
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.predictor = SalesPredictor(file_path) # Crear una instancia del predictor con el archivo seleccionado
            self.predictor.train_model() # Entrenar el modelo de predicción
            self.check_low_stock() # Verificar productos con stock bajo
            self.predictor.add_low_stock_warning() # Añadir advertencias de stock bajo
            self.predict_button.config(state=tk.NORMAL) # Habilitar el botón de predicción
            messagebox.showinfo("Información", "Archivo cargado correctamente")

    def check_low_stock(self):
        # Filtrar los productos con unidades menores a 1000
        low_stock = self.predictor.data[self.predictor.data['Unidades'] < 1000]
        if not low_stock.empty:
            num_low_stock_products = len(low_stock)
            self.show_alert(f"Hay {num_low_stock_products} productos con stock bajo", "warning")
    
    def show_alert(self, message, alert_type):
        # Crear una ventana emergente de alerta
        alert_window = tk.Toplevel(self.root)
        alert_window.title("Alerta")
        alert_window.geometry("300x150")
        alert_window.configure(bg='white')

        if alert_type == "warning":
            alert_window.configure(bg='red')

        label = tk.Label(alert_window, text=message, font=('Arial', 12), bg=alert_window.cget('bg'), fg='white' if alert_type == "warning" else 'black')
        label.pack(expand=True, fill='both', padx=20, pady=20)

        close_button = tk.Button(alert_window, text="Cerrar", command=alert_window.destroy, font=('Arial', 12))
        close_button.pack(pady=10)

        alert_window.transient(self.root)
        alert_window.grab_set()
        self.root.wait_window(alert_window)

    def predict_sales(self):
        # Ejecutar las funciones para encontrar productos principales, graficar predicciones y generar el reporte
        self.find_top_products()
        self.plot_predictions()
        self.generate_report()
        self.save_pdf_report()

    def find_top_products(self):
        # Encontrar los productos más vendidos
        df = self.predictor.data
        monthly_sales = df.groupby(['Mes', 'Tipo de producto'])['Unidades'].sum().reset_index()
        top_products = monthly_sales.groupby('Tipo de producto')['Unidades'].sum().nlargest(3).index.tolist()
        self.top_products = top_products

    def plot_predictions(self):
        # Graficar las predicciones de ventas
        df = self.predictor.data
        num_products = len(self.top_products)
        num_cols = 2
        num_rows = (num_products * 2 + num_cols - 1) // num_cols
        
        plt.figure(figsize=(15, num_rows * 5))

        colors = ['blue', 'green', 'red', 'purple', 'orange', 'cyan']
        pred_color = 'orange'  # Color para las predicciones

        for i, product in enumerate(self.top_products):
            product_data = df[df['Tipo de producto'] == product]
            monthly_sales = product_data.groupby('Mes')['Unidades'].sum().reindex(range(1, 13), fill_value=0)

            plt.subplot(num_rows, num_cols, 2*i + 1)
            plt.bar(monthly_sales.index, monthly_sales.values, color='black', alpha=0.7)
            plt.title(f'Ventas Reales de {product}', fontsize=9)
            plt.xlabel('Mes', fontsize=9)
            plt.ylabel('Unidades Vendidas', fontsize=9)
            plt.xticks(range(1, 13), [f'Mes {i}' for i in range(1, 13)], rotation=40, fontsize=8)
            plt.grid(True, linestyle='--', alpha=0.5)

            product_x_test = self.predictor.X_test[self.predictor.X_test['Tipo de producto'] == product]
            y_pred = self.predictor.predict_for_product(product_x_test)
            predicted_sales = pd.Series(y_pred).groupby(self.predictor.X_test[self.predictor.X_test['Tipo de producto'] == product].index // len(product_x_test)).sum()
            
            # Resaltar si hay aumento en las ventas predichas
            increased_sales = predicted_sales[:12] > monthly_sales.values[:12]

            plt.subplot(num_rows, num_cols, 2*i + 2)
            bars = plt.bar(monthly_sales.index, predicted_sales[:12], color=pred_color, alpha=0.7)
            
            # Cambiar el color de las barras que muestran un aumento
            for bar, increase in zip(bars, increased_sales):
                if increase:
                    bar.set_color('red')  # Color para los aumentos

            plt.title(f'Predicción de Ventas de {product}', fontsize=9)
            plt.xlabel('Mes', fontsize=9)
            plt.ylabel('Unidades Vendidas', fontsize=9)
            plt.xticks(range(1, 13), [f'Mes {i}' for i in range(1, 13)], rotation=40, fontsize=8)
            plt.grid(True, linestyle='--', alpha=0.5)

        plt.tight_layout()
        plt.show()

    def generate_report(self):
        # Generar el reporte de ventas mensual y actualizar el Arbol de vistas
        df = self.predictor.data
        monthly_sales = df.groupby(['Mes', 'Tipo de producto'])['Unidades'].sum().reset_index()

        for row in self.tree.get_children():
            self.tree.delete(row)

        for _, row in monthly_sales.iterrows():
            self.tree.insert('', 'end', values=(row['Mes'], row['Tipo de producto'], row['Unidades']))

    def save_pdf_report(self):
        # Guardar el reporte generado en un archivo PDF
        pdf = FPDF()
        pdf.add_page()

        pdf.set_font("Arial", size=12)
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Reporte de Predicción de Ventas Anual", ln=True, align='C')
        pdf.ln(10)
        
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(40, 10, "Mes", border=1, align='C')
        pdf.cell(80, 10, "Tipo de Producto", border=1, align='C')
        pdf.cell(40, 10, "Unidades", border=1, align='C')
        pdf.ln()

        pdf.set_font("Arial", size=12)
        
        for row_id in self.tree.get_children():
            row = self.tree.item(row_id)['values']
            pdf.cell(40, 10, str(row[0]), border=1, align='C')
            pdf.cell(80, 10, str(row[1]), border=1, align='C')
            pdf.cell(40, 10, str(row[2]), border=1, align='C')
            pdf.ln()

        pdf.output("prediccion_ventas.pdf")

    def voice_search(self):
        # Iniciar la búsqueda por voz
        recognizer = sr.Recognizer()
        self.result_label.config(text="Grabando...")
        self.root.update()  # Asegúrate de que la interfaz gráfica se actualice

        with sr.Microphone() as source:
            # Captura el audio del micrófono
            audio = recognizer.listen(source)

        try:
            # Intenta reconocer el texto del audio en español
            text = recognizer.recognize_google(audio, language="es-ES")
            self.result_label.config(text=f"Buscando: {text}")
            
            # Dependiendo del texto reconocido, llama a la función correspondiente
            if "producto con el stock más bajo" in text:
                self.show_low_stock_products()
            elif "unidades quedan del producto" in text:
                product = text.split("del producto")[-1].strip()
                self.show_stock_for_product(product)
            elif "productos tienen stock bajo este mes" in text:
                self.show_low_stock_this_month()
            elif "productos más vendidos" in text:
                self.show_top_selling_products()
            elif "reporte de ventas del mes pasado" in text:
                self.generate_monthly_sales_report()
            elif "productos se vendieron en total" in text:
                self.show_total_sales()
            elif "productos con el stock bajo" in text:
                self.show_low_stock_products_total()
            else:
                self.result_label.config(text="Pregunta no reconocida")
            
        except sr.UnknownValueError:
            # Error en caso de que no se entienda el audio
            self.result_label.config(text="No se entendió el audio")
        except sr.RequestError:
            # Error en caso de fallo en la conexión con el servicio de reconocimiento
            self.result_label.config(text="Error al conectar con el servicio de reconocimiento")

    def show_low_stock_products(self):
        # Muestra los productos con stock bajo (menos de 10 unidades) en la interfaz
        low_stock = self.predictor.data[self.predictor.data['Unidades'] < 10]
        self.update_treeview(low_stock)

    def show_stock_for_product(self, product):
        # Muestra el stock de un producto específico, buscando el nombre más cercano si no se encuentra
        product_names = self.predictor.product_names.tolist()
        if product not in product_names:
            closest_match = process.extractOne(product, product_names)
            if closest_match and closest_match[1] > 80:  # Un umbral de coincidencia del 80%
                product = closest_match[0]
            else:
                self.result_label.config(text="Producto no encontrado")
                return
        
        product_code = product_names.index(product)
        stock = self.predictor.data[self.predictor.data['Tipo de producto'] == product_code]
        self.update_treeview(stock)

    def show_low_stock_this_month(self):
        # Muestra los productos con stock bajo (menos de 1000 unidades) en el mes actual
        current_month = pd.to_datetime('now').month
        low_stock = self.predictor.data[(self.predictor.data['Mes'] == current_month) & (self.predictor.data['Unidades'] < 1000)]
        low_stock = low_stock.sort_values(by='Unidades').head(5)
        self.update_treeview(low_stock)

    def show_top_selling_products(self):
        # Muestra los 3 productos más vendidos
        top_products = self.predictor.data.groupby('Tipo de producto')['Unidades'].sum().nlargest(3).index
        top_selling = self.predictor.data[self.predictor.data['Tipo de producto'].isin(top_products)]
        self.update_treeview(top_selling)

    def generate_monthly_sales_report(self):
        # Genera un reporte mensual de ventas y muestra el resultado
        report_path = self.predictor.generate_monthly_sales_report()
        self.result_label.config(text=f"Reporte mensual generado: {report_path}")
        messagebox.showinfo("Información", f"Reporte mensual generado: {report_path}")

    def show_total_sales(self):
        # Muestra las ventas totales
        total_sales = self.predictor.data['Unidades'].sum()
        self.result_label.config(text=f"Ventas totales: {total_sales}")

    def update_treeview(self, data):
        # Actualiza el Arbol con los datos proporcionados
        for row_id in self.tree.get_children():
            self.tree.delete(row_id)
        for _, row in data.iterrows():
            self.tree.insert('', 'end', values=(row['Mes'], row['Tipo de producto'], row['Unidades']))

    def plot_sales_predictions(self, product, predictions):
        # Implementa la visualización de predicciones (POR DEFINIR)
        pass
    
    def show_low_stock_products_total(self):
        # Muestra los productos con stock bajo (menos de 1000 unidades) en general
        low_stock = self.predictor.data[self.predictor.data['Unidades'] < 1000]
        self.update_treeview(low_stock)

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesPredictionApp(root)
    root.mainloop()
