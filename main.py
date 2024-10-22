import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
from tkinter import Tk

def main():
    # Selección del archivo Excel
    root = Tk()
    root.withdraw()  # Ocultar la ventana de Tkinter
    file_path = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        print("No se seleccionó ningún archivo.")
        return

    # Cargar el archivo Excel
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return

    # Asegurarse de que las columnas necesarias estén presentes
    required_columns = ["x", "y"]
    for col in required_columns:
        if col not in df.columns:
            print(f"La columna '{col}' no está presente en el archivo Excel.")
            return

    x = df['x']
    y = df['y']

    # Cálculo de Sxx, Syy, Sxy
    n = len(x)
    Sxx = np.sum(x ** 2) - (np.sum(x) ** 2) / n
    Syy = np.sum(y ** 2) - (np.sum(y) ** 2) / n
    Sxy = np.sum(x * y) - (np.sum(x) * np.sum(y)) / n

    # Cálculo del coeficiente de correlación (r)
    r = Sxy / np.sqrt(Sxx * Syy)
    print(f"Coeficiente de correlación (r): {r}")

    # Cálculo del coeficiente de determinación (R^2)
    R_squared = (r ** 2)
    print(f"Coeficiente de determinación (R^2): {R_squared}")

    # Gráfica de dispersión con elementos mejorados
    plt.scatter(x, y, color='blue', label='Datos')
    plt.xlabel('Publicidad (minutos)', fontsize=12)
    plt.ylabel('Ventas (unidades)', fontsize=12)
    plt.title('Relación entre Publicidad y Ventas', fontsize=14, fontweight='bold')
    plt.grid(True, linestyle='--', alpha=0.7)

    # Agregar etiquetas a los puntos de datos
    for i in range(len(x)):
        plt.text(x[i], y[i], f'({x[i]}, {y[i]})', fontsize=8, ha='right')

    # Agregar línea de mejor ajuste
    m, b = np.polyfit(x, y, 1)
    plt.plot(x, m * x + b, color='red', linestyle='-', linewidth=2, label='Línea de mejor ajuste')

    plt.legend()
    plt.show()

if __name__ == "__main__":
    main()
