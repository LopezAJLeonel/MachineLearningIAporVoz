# Proyecto: Inventario con Inteligencia Artificial
# Responsable: Jose Leonel Lopez Almeida
# Versión: 1.0
# # Version: 1.1 Actualizada

# Fecha_Creación: 06/06/2024
# Fecha_Modificación: 10/08/2024
# Descripción: Este script se encarga de predecir ventas y generar reportes de advertencia de stock bajo.
# Utiliza un modelo de regresión lineal para hacer las predicciones basado en los datos históricos de ventas.
# Adicionalmente, permite generar un reporte mensual de ventas de un producto.

import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error
from sklearn.preprocessing import LabelEncoder

class SalesPredictor:
    def __init__(self, file_path):
        
        self.data = pd.read_excel(file_path)  # Carga los datos desde el archivo Excel proporcionado.
        self.file_path = file_path  # Guarda la ruta del archivo original.
        self.prepare_data()  # Prepara los datos para el modelo.
        self.model = LinearRegression()  # Inicializa el modelo de regresión lineal.

    def prepare_data(self):
        
        self.data['Fecha pedido'] = pd.to_datetime(self.data['Fecha pedido'])  # Convierte la columna de fechas a formato datetime.
        self.data['Mes'] = self.data['Fecha pedido'].dt.month  # Extrae el mes de la fecha.
        self.data['Año'] = self.data['Fecha pedido'].dt.year  # Extrae el año de la fecha.
        
        # Codifica la variable 'Tipo de producto' de texto a valores numéricos.
        le = LabelEncoder()
        self.data['Tipo de producto'] = le.fit_transform(self.data['Tipo de producto'])
        self.product_names = le.classes_  # Guarda los nombres originales de los productos.

        # Define las variables independientes (X) y la variable dependiente (y).
        self.X = self.data[['Mes', 'Año', 'Tipo de producto']]
        self.y = self.data['Unidades']

        # Divide los datos en conjuntos de entrenamiento y prueba.
        self.X_train, self.X_test, self.y_train, self.y_test = train_test_split(self.X, self.y, test_size=0.2, random_state=42)

    def train_model(self):
        # Entrena el modelo de regresión lineal usando los datos de entrenamiento.
        self.model.fit(self.X_train, self.y_train)  # Ajusta el modelo a los datos de entrenamiento.

    def predict(self):
        
        y_pred = self.model.predict(self.X_test)  # Predice las unidades usando los datos de prueba.
        mse = mean_squared_error(self.y_test, y_pred)  # Calcula el error cuadrático medio entre las predicciones y los valores reales.
        return y_pred, mse

    def predict_for_product(self, X_test_product):
        return self.model.predict(X_test_product)  # Retorna las predicciones para los datos proporcionados.

    def add_low_stock_warning(self):
        # Define el umbral de stock bajo.
        low_stock_threshold = 1000
        
        # Añade una columna de advertencia si las unidades están por debajo del umbral.
        self.data['Advertencia'] = self.data['Unidades'].apply(lambda x: 'Stock Bajo' if x < low_stock_threshold else 'Stock Adecuado')
        
        # Selecciona solo las columnas necesarias para el reporte.
        columns_to_keep = ['Mes', 'Tipo de producto', 'Unidades', 'Advertencia']
        low_stock_report = self.data[columns_to_keep]
        
        # Guarda el reporte en un archivo Excel.
        low_stock_report.to_excel("reporte_stock_bajo.xlsx", index=False)

    def generate_monthly_sales_report(self):
        # Agrupa los datos por tipo de producto y calcula el total de unidades vendidas.
        monthly_sales = self.data.groupby('Tipo de producto')['Unidades'].sum().reset_index()
    
        # Guarda el reporte en un archivo Excel.
        report_path = "reporte_ventas_mensuales.xlsx"
        monthly_sales.to_excel(report_path, index=False)
    
        return report_path  # Retorna la ruta del archivo generado.

    def get_product_names(self):
        return self.product_names  # Retorna los nombres originales de los productos.
