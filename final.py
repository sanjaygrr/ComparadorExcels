import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox


def extraer_codigos_autorizacion(description):
    pattern = r'c[oó]digo de autorizaci[oó]n[.:_]?[\s]?([a-zA-Z0-9]+)'
    return re.findall(pattern, description, re.IGNORECASE)


def cargar_archivo(tipo):
    filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    if file_path:
        if tipo == 'fdm':
            global fdm_path
            fdm_path = file_path
            label_fdm.config(
                text=f"Archivo FDM cargado: {file_path.split('/')[-1]}", fg="green")
        elif tipo == 'transbank':
            global transbank_path
            transbank_path = file_path
            label_transbank.config(
                text=f"Archivo Transbank cargado: {file_path.split('/')[-1]}", fg="green")


def leer_archivo(path):
    if path.endswith('.csv'):
        return pd.read_csv(path)
    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    else:
        raise ValueError("Formato de archivo no soportado")


def procesar_archivos():
    try:
        fdm_df = leer_archivo(fdm_path)
        transbank_df = leer_archivo(transbank_path)

        transbank_df = transbank_df.loc[:, ~
                                        transbank_df.columns.str.contains('^Unnamed')]

        fdm_credit_card = fdm_df[fdm_df['Account'] == 'Credit Card'].copy()
        fdm_credit_card['Codigos_Autorizacion'] = fdm_credit_card['Description'].apply(
            lambda x: extraer_codigos_autorizacion(x) if isinstance(x, str) else [])

        fdm_expanded = fdm_credit_card.explode('Codigos_Autorizacion')

        codigos_transbank = set(transbank_df['C�digo Autorizaci�n'])
        transacciones_coincidentes = fdm_expanded[fdm_expanded['Codigos_Autorizacion'].isin(
            codigos_transbank)]
        transacciones_faltantes_fdm = fdm_expanded[~fdm_expanded['Codigos_Autorizacion'].isin(
            codigos_transbank)]
        transacciones_faltantes_transbank = transbank_df[~transbank_df['C�digo Autorizaci�n'].isin(
            fdm_expanded['Codigos_Autorizacion'])]

        with pd.ExcelWriter('transacciones_comparativas_ajustadas.xlsx') as writer:
            transacciones_coincidentes.to_excel(
                writer, sheet_name='Coincidentes', index=False)
            transacciones_faltantes_fdm.to_excel(
                writer, sheet_name='Faltantes en FDM', index=False)
            transacciones_faltantes_transbank.to_excel(
                writer, sheet_name='Faltantes en Transbank', index=False)

        messagebox.showinfo(
            "Éxito", "El archivo Excel ha sido creado con éxito.")
    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title("Comparador de Archivos Excel")
root.geometry("700x450")

frame = tk.Frame(root, bg="#F0F0F0")
frame.pack(pady=20)

label_intro = tk.Label(frame, text="Seleccione los archivos CSV o Excel para comparar", font=(
    "Helvetica", 16), bg="#F0F0F0")
label_intro.grid(row=0, columnspan=2, padx=10, pady=10)

label_carga = tk.Label(frame, text="Cargar archivos",
                       font=("Helvetica", 14), bg="#F0F0F0")
label_carga.grid(row=1, columnspan=2, padx=10, pady=10)

label_fdm = tk.Label(frame, text="No se ha seleccionado archivo FDM", font=(
    "Helvetica", 12), bg="#F0F0F0", fg="black")
label_fdm.grid(row=2, column=0, padx=10, pady=5)
boton_fdm = tk.Button(frame, text="Cargar FDM", command=lambda: cargar_archivo(
    'fdm'), bg="#007ACC", fg="white")
boton_fdm.grid(row=2, column=1, padx=10, pady=5)

label_transbank = tk.Label(frame, text="No se ha seleccionado archivo Transbank", font=(
    "Helvetica", 12), bg="#F0F0F0", fg="black")
label_transbank.grid(row=3, column=0, padx=10, pady=5)
boton_transbank = tk.Button(frame, text="Cargar Transbank", command=lambda: cargar_archivo(
    'transbank'), bg="#007ACC", fg="white")
boton_transbank.grid(row=3, column=1, padx=10, pady=5)

boton_procesar = tk.Button(root, text="Procesar Archivos",
                           command=procesar_archivos, bg="#4CAF50", fg="white", font=("Helvetica", 16))
boton_procesar.pack(pady=20)

root.mainloop()
