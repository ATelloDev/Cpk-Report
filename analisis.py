import tkinter as tk
from tkinter import filedialog
import os
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from PIL import Image

# Define los estilos solo si aún no están definidos
styles = getSampleStyleSheet()
if 'Heading1' not in styles:
    styles.add(ParagraphStyle(name='Heading1', fontSize=14, textColor=colors.blue, spaceBefore=12))
    styles.add(ParagraphStyle(name='Heading2', fontSize=12, textColor=colors.black, spaceBefore=12))
    styles.add(ParagraphStyle(name='Normal', fontSize=10, textColor=colors.black))

def analizar_archivos():
    global folder_selected
    global archivos
    global objetivo
    global lsl
    global usl
    
    objetivo = float(entry_objetivo.get())
    lsl = float(entry_lsl.get())
    usl = float(entry_usl.get())
    
    folder_selected = filedialog.askdirectory()
    archivos = [os.path.join(folder_selected, archivo) for archivo in os.listdir(folder_selected) if archivo.endswith('.csv')]
    
    btn_generar_reporte.config(state=tk.NORMAL)

def generar_reporte():
    pdf = SimpleDocTemplate("Informe.pdf", pagesize=letter)
    story = []
    
    for archivo in archivos:
        contenido = analisis_proceso(archivo, objetivo, lsl, usl)
        story.extend(contenido)

        # Agregar el gráfico al informe
        agregar_graficos_a_reporte("Informe.pdf", f"grafico_{os.path.basename(archivo)}.png", f"grafico_control_{os.path.basename(archivo)}.png")
        
    pdf.build(story)

def analisis_proceso(archivo, objetivo, lsl, usl):
    story = []
    
    try:
        datos = pd.read_csv(archivo)['muestras'].values
    except KeyError:
        story.append(Paragraph(f"La columna 'muestras' no existe en el archivo {archivo}.", styles['Normal']))
        return story

    sample_mean, sample_std_dev, median, rango, cuartiles = calcular_estadisticas(datos)
    sample_mean_cpk, num_muestra, sample_std_dev_short_term, sample_std_dev_long_term, cp, cpl, cpu, cpk, pp, ppl, ppu, ppk = calcular_cpk(
        datos, objetivo, lsl, usl)
    lsc = sample_mean_cpk + 3 * sample_std_dev_short_term
    lic = sample_mean_cpk - 3 * sample_std_dev_short_term

    # Agregar estadísticas descriptivas al informe
    story.append(Paragraph(f"Estadísticas Descriptivas - {os.path.basename(archivo)}", styles['Heading1']))
    tabla_estadisticas = [
        ["Media de la muestra:", sample_mean],
        ["Desviación estándar de la muestra:", sample_std_dev],
        ["Mediana:", median],
        ["Rango:", rango],
        ["Cuartiles:", f"Q1={cuartiles[0]}, Q2={cuartiles[1]}, Q3={cuartiles[2]}"]
    ]
    story.append(Table(tabla_estadisticas))
    story.append(Spacer(1, 12))

    # Crear y guardar el histograma de los datos
    plt.hist(datos, bins=20, color='skyblue', edgecolor='black', alpha=0.7)
    plt.xlabel('Valor')
    plt.ylabel('Frecuencia')
    plt.title('Histograma de Muestras')
    grafico_filename = f"grafico_{os.path.basename(archivo)}.png"
    plt.savefig(grafico_filename)
    plt.close()

    # Crear y guardar el gráfico de control
    np.random.seed(0)
    datos_control = np.random.normal(sample_mean, sample_std_dev, 100)
    plt.plot(datos_control, color='green')
    plt.axhline(y=usl, color='r', linestyle='--', label='USL')
    plt.axhline(y=lsl, color='r', linestyle='--', label='LSL')
    plt.xlabel('Muestra')
    plt.ylabel('Valor')
    plt.title('Gráfico de Control')
    plt.legend()
    grafico_control_filename = f"grafico_control_{os.path.basename(archivo)}.png"
    plt.savefig(grafico_control_filename)
    plt.close()



    # Agregar el histograma al informe
    story.append(Paragraph("Histograma de Muestras:", styles['Heading2']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<img src='{grafico_filename}' width='400' height='300'/>", styles['Normal']))
    story.append(Spacer(1, 24))  # Espacio adicional después del histograma

    # Agregar el gráfico de control al informe
    story.append(Paragraph("Gráfico de Control:", styles['Heading2']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<img src='{grafico_control_filename}' width='400' height='300'/>", styles['Normal']))
    story.append(Spacer(1, 24))  # Espacio adicional después del gráfico de control

    # Agregar resultados de CPK al informe
    story.append(Paragraph("Estadísticas de Capacidad del Proceso:", styles['Heading2']))
    estadisticas_cpk = [
        ["Media de la muestra (CPK):", sample_mean_cpk],
        ["Número de muestras:", num_muestra],
        ["Desviación estándar a corto plazo:", sample_std_dev_short_term],
        ["Desviación estándar a largo plazo:", sample_std_dev_long_term],
        ["CP del proceso:", cp],
        ["CPL del proceso:", cpl],
        ["CPU del proceso:", cpu],
        ["Cpk del proceso:", cpk],
        ["Pp del proceso:", pp],
        ["PPL del proceso:", ppl],
        ["PPU del proceso:", ppu],
        ["Ppk del proceso:", ppk],
        ["\n" * 3]
    ]

    story.append(Table(estadisticas_cpk))
    story.append(Spacer(1, 12))

    return story

def calcular_estadisticas(datos):
    sample_mean = np.mean(datos)
    sample_std_dev = np.std(datos, ddof=1)
    median = np.median(datos)
    rango = np.ptp(datos)
    cuartiles = np.percentile(datos, [25, 50, 75])
    return sample_mean, sample_std_dev, median, rango, cuartiles

def calcular_cpk(datos, objetivo, lsl, usl):
    sample_mean = np.mean(datos)
    num_muestra = len(datos)
    sample_std_dev_short_term = np.std(datos, ddof=1)
    sample_std_dev_long_term = np.std(datos)

    cp = (usl - lsl) / (6 * sample_std_dev_short_term)
    cpl = (sample_mean - lsl) / (3 * sample_std_dev_short_term)
    cpu = (usl - sample_mean) / (3 * sample_std_dev_short_term)
    cpk = min(cpl, cpu)

    pp = (usl - lsl) / (6 * sample_std_dev_long_term)
    ppl = (sample_mean - lsl) /(3 * sample_std_dev_long_term)
    ppu = (usl - sample_mean) / (3 * sample_std_dev_long_term)
    ppk = min(ppl, ppu)

    return sample_mean, num_muestra, sample_std_dev_short_term, sample_std_dev_long_term, cp, cpl, cpu, cpk, pp, ppl, ppu, ppk

def agregar_graficos_a_reporte(pdf_filename, histograma_filename, control_filename):
    c = canvas.Canvas(pdf_filename, pagesize=letter)
    c.drawString(100, 750, "Reporte de Ventas")

    # Insertar el histograma de muestras en el informe PDF
    c.drawImage(histograma_filename, 100, 600, width=400, height=300)

    # Insertar el gráfico de control en el informe PDF
    c.drawImage(control_filename, 100, 300, width=400, height=300)

    c.save()

# Configuración de la ventana principal
root = tk.Tk()
root.title("Análisis de Proceso")

# Campos de entrada
label_objetivo = tk.Label(root, text="Objetivo:")
label_objetivo.grid(row=0, column=0, padx=5, pady=5)
entry_objetivo = tk.Entry(root)
entry_objetivo.grid(row=0, column=1, padx=5, pady=5)

label_lsl = tk.Label(root, text="LSL:")
label_lsl.grid(row=1, column=0, padx=5, pady=5)
entry_lsl = tk.Entry(root)
entry_lsl.grid(row=1, column=1, padx=5, pady=5)

label_usl = tk.Label(root, text="USL:")
label_usl.grid(row=2, column=0, padx=5, pady=5)
entry_usl = tk.Entry(root)
entry_usl.grid(row=2, column=1, padx=5, pady=5)

# Botón para analizar archivos
btn_analizar = tk.Button(root, text="Seleccionar Carpeta", command=analizar_archivos)
btn_analizar.grid(row=3, column=0, padx=5, pady=5)

# Botón para generar el informe
btn_generar_reporte = tk.Button(root, text="Generar Reporte", command=generar_reporte, state=tk.DISABLED)
btn_generar_reporte.grid(row=3, column=1, padx=5, pady=5)

root.mainloop()

