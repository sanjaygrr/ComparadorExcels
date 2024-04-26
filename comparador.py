import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to extract authorization codes from descriptions
def extract_codes(description):
    pattern = re.compile(r'\b[cC][aA][: ]?\d+')
    matches = pattern.findall(description)
    codes = [re.search(r'\d+', match).group() for match in matches]
    return codes

# Function to process the Excel file
def process_file(file_path):
    try:
        # Read the Excel file
        xl = pd.ExcelFile(file_path)
        fdm_data = xl.parse(sheet_name='FDM')
        transbank_data = xl.parse(sheet_name='Transbank')

        # Extract authorization codes
        fdm_data['Códigos Extraídos'] = fdm_data['Description'].astype(str).apply(extract_codes)
        
        # Flatten the list of codes and get unique values
        fdm_codes = fdm_data['Códigos Extraídos'].explode().dropna().unique()
        transbank_codes = transbank_data['Código Autorización'].astype(str).unique()

        # Find matches and mismatches
        matches = set(fdm_codes) & set(transbank_codes)
        mismatches_fdm = set(fdm_codes) - set(transbank_codes)
        mismatches_transbank = set(transbank_codes) - set(fdm_codes)

        # Generate the data for the matched codes sheet
        datos_coincidentes = generate_matched_data(matches, fdm_data, transbank_data)

        # Generate the data for the FDM mismatches sheet
        datos_no_coincidentes_fdm = fdm_data[fdm_data['Códigos Extraídos'].apply(lambda x: bool(set(x) & mismatches_fdm))]

        # Generate the data for the Transbank mismatches sheet
        datos_no_coincidentes_transbank = transbank_data[transbank_data['Código Autorización'].isin(mismatches_transbank)]

        # Ask the user for a location and filename to save the new Excel file
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Guardar archivo como"
        )

        # Write the data to the new Excel file
        if output_path:
            with pd.ExcelWriter(output_path) as writer:
                datos_coincidentes.to_excel(writer, sheet_name='Coincidencias', index=False)
                datos_no_coincidentes_fdm.to_excel(writer, sheet_name='No Coincidentes FDM', index=False)
                datos_no_coincidentes_transbank.to_excel(writer, sheet_name='No Coincidentes Transbank', index=False)
            messagebox.showinfo("Éxito", "El archivo ha sido procesado con éxito!")
        else:
            messagebox.showwarning("Advertencia", "No se seleccionó ubicación para guardar.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to retrieve rows with matching codes and construct the final DataFrame
def generate_matched_data(matches, fdm_data, transbank_data):
    matching_details_data = []
    for code in matches:
        fdm_rows = fdm_data[fdm_data['Códigos Extraídos'].apply(lambda codes: code in codes)]
        transbank_row = transbank_data[transbank_data['Código Autorización'].astype(str) == code].iloc[0]
        for _, fdm_row in fdm_rows.iterrows():
            matching_details_data.append({
                'Código Autorización': code,
                'Usuario FDM': fdm_row['User'],
                'Monto FDM': fdm_row['Amount'],
                'Monto Transbank': transbank_row['Monto Afecto'],
                'Fecha y Hora FDM': fdm_row['Date Time'],
                'Fecha Venta Transbank': transbank_row['Fecha Venta'],
                'Categoría FDM': fdm_row['Category'],
                'Moneda FDM': fdm_row['Currency'],
                'Huésped FDM': fdm_row['Guest'],
                'Movimiento Transbank': transbank_row['Tipo Movimiento'],
                'Tarjeta Transbank': transbank_row['Tipo Tarjeta']
            })
    return pd.DataFrame(matching_details_data)

# Function to open a dialog and select a file
def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")],
        title="Seleccionar archivo Excel"
    )
    if file_path:
        process_file(file_path)

# Set up the GUI window
root = tk.Tk()
root.title("Procesador de Excel")
root.geometry("400x200")

# Create a button to select the Excel file
open_button = tk.Button(root, text="Seleccionar Archivo Excel", command=select_file, height=2, width=20)
open_button.pack(pady=20)

# Start the GUI loop
root.mainloop()
