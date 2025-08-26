import customtkinter as ctk
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
import pythoncom
from win32com.shell import shell
import re

# -----------------------------
# Función para resolver accesos directos (.lnk)
# -----------------------------
def resolve_lnk(lnk_path):
    if not lnk_path.lower().endswith('.lnk'):
        return lnk_path
    shell_link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None,
        pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
    )
    persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Load(lnk_path)
    path, _ = shell_link.GetPath(shell.SLGP_UNCPRIORITY)
    return path

# -----------------------------
# Diccionario de traducción
# -----------------------------
LANG_DICT = {
    "ES": {
        "title": "Procesador Excel HOP",
        "select_button": "Seleccionar archivos",
        "choose_folder": "Elegir carpeta de salida",
        "process": "Procesar archivos",
        "files_selected": "Archivos seleccionados:",
        "files_processed": "Archivos procesados:"
    },
    "EN": {
        "title": "HOP Excel Processor",
        "select_button": "Select files",
        "choose_folder": "Choose output folder",
        "process": "Process files",
        "files_selected": "Selected files:",
        "files_processed": "Processed files:"
    }
}

# -----------------------------
# Clase principal de la app
# -----------------------------
class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ----------------------------- Icono de ventana -----------------------------
        self.iconbitmap("logo.ico")  # <-- Ícono para ventana y barra de tareas

        self.geometry("900x650")
        self.title(LANG_DICT["ES"]["title"])
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.current_lang = "ES"
        self.selected_files = []
        self.output_folder = os.getcwd()

        # ----------------------------- Contenedor central -----------------------------
        self.center_frame = ctk.CTkFrame(self, fg_color="#f0f0f0", corner_radius=20)
        self.center_frame.pack(pady=10, padx=30, fill="both", expand=True)

        # ----------------------------- Menú de idioma -----------------------------
        self.lang_option = ctk.CTkOptionMenu(
            self, values=["ES","EN"], command=self.change_language, width=120
        )
        self.lang_option.place(relx=0.95, y=10, anchor="ne")
        self.lang_option.tkraise()

        # ----------------------------- Botón seleccionar archivos -----------------------------
        self.select_button = ctk.CTkButton(
            self.center_frame,
            text=LANG_DICT[self.current_lang]["select_button"],
            command=self.select_files,
            corner_radius=15,
            width=220
        )
        self.select_button.pack(pady=(20,10))

        # ----------------------------- Frame contenedor para archivos seleccionados con scroll -----------------------------
        self.selected_files_container = ctk.CTkFrame(self.center_frame, height=120, corner_radius=15)
        self.selected_files_container.pack(fill="x", padx=20, pady=5)
        self.selected_files_container.pack_propagate(False)

        self.selected_files_frame = ctk.CTkScrollableFrame(self.selected_files_container)
        self.selected_files_frame.pack(fill="both", expand=True)
        self.selected_files_frame.grid_columnconfigure(0, weight=1)

        # ----------------------------- Checkboxes para opciones -----------------------------
        self.format_no_weights_var = ctk.IntVar(value=1)
        self.calc_pesindiv_var = ctk.IntVar()

        self.format_no_weights_cb = ctk.CTkCheckBox(
            self.center_frame,
            text="Formato sin pesos estimados",
            variable=self.format_no_weights_var
        )
        self.format_no_weights_cb.pack(pady=(10,0))
        self.format_no_weights_desc = ctk.CTkLabel(
            self.center_frame,
            text='Pesos de dos individuos aparecen como NA\ncolumna "Tipus" cambiada al diccionario de MRAG.',
            fg_color="#d0e6ff",
            text_color="#003366",
            corner_radius=10,
            anchor="center",
            justify="center"
        )
        self.format_no_weights_desc.pack(padx=20, pady=(2,10))
        self.format_no_weights_desc.configure(wraplength=400)

        self.calc_pesindiv_cb = ctk.CTkCheckBox(
            self.center_frame,
            text="Calcular PesIndiv y normalizar Tipus",
            variable=self.calc_pesindiv_var
        )
        self.calc_pesindiv_cb.pack(pady=(10,0))
        self.calc_pesindiv_desc = ctk.CTkLabel(
            self.center_frame,
            text="Calcula Peso Individual de cada pez en base a la relación de largo y ancho y la suma total del Pes M.\nNo influye en el peso total del HOP.",
            fg_color="#d0e6ff",
            text_color="#003366",
            corner_radius=10,
            anchor="center",
            justify="center"
        )
        self.calc_pesindiv_desc.pack(padx=20, pady=(2,10))
        self.calc_pesindiv_desc.configure(wraplength=400)

        # ----------------------------- Botón elegir carpeta de salida -----------------------------
        self.folder_button = ctk.CTkButton(
            self.center_frame,
            text=LANG_DICT[self.current_lang]["choose_folder"],
            command=self.select_output_folder,
            corner_radius=15,
            width=220
        )
        self.folder_button.pack(pady=10)

        # ----------------------------- Label para mostrar carpeta de salida -----------------------------
        self.output_folder_label = ctk.CTkLabel(
            self.center_frame,
            text=f"Carpeta de salida: {self.output_folder}",
            anchor="w",
            justify="left",
            wraplength=850  # Ajustable para que no se corte
        )
        self.output_folder_label.pack(padx=20, pady=(2,10), fill="x")

        # ----------------------------- Botón procesar archivos -----------------------------
        self.process_button = ctk.CTkButton(
            self.center_frame,
            text=LANG_DICT[self.current_lang]["process"],
            command=self.procesar_archivos,
            corner_radius=15,
            width=220
        )
        self.process_button.pack(pady=10)

        # ----------------------------- Frame contenedor para archivos procesados con scroll -----------------------------
        self.processed_files_container = ctk.CTkFrame(self.center_frame, height=180, corner_radius=15)
        self.processed_files_container.pack(fill="x", padx=20, pady=10)
        self.processed_files_container.pack_propagate(False)

        self.processed_files_frame = ctk.CTkScrollableFrame(self.processed_files_container)
        self.processed_files_frame.pack(fill="both", expand=True)
        self.processed_files_frame.grid_columnconfigure(0, weight=1)

        # ----------------------------- Firma de la app -----------------------------
        self.signature_label = ctk.CTkLabel(
            self,
            text="Creado por Bernardo R. para los observadores de MRAG en la granja de L´Ametlla de Mar",
            text_color="#666666",
            font=("Arial", 9)
        )
        self.signature_label.pack(side="bottom", pady=10)

    # ----------------------------- Cambiar idioma -----------------------------
    def change_language(self, choice):
        self.current_lang = choice
        self.title(LANG_DICT[choice]["title"])
        self.select_button.configure(text=LANG_DICT[choice]["select_button"])
        self.folder_button.configure(text=LANG_DICT[choice]["choose_folder"])
        self.process_button.configure(text=LANG_DICT[choice]["process"])

    # ----------------------------- Seleccionar archivos -----------------------------
    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        for path in paths:
            real_path = resolve_lnk(path)
            if real_path not in self.selected_files:
                self.selected_files.append(real_path)
        self.update_selected_files_label()

    # ----------------------------- Actualizar lista de archivos seleccionados -----------------------------
    def update_selected_files_label(self):
        for widget in self.selected_files_frame.winfo_children():
            widget.destroy()

        for i, f in enumerate(self.selected_files):
            file_label = ctk.CTkLabel(self.selected_files_frame, text=os.path.basename(f), anchor="w")
            file_label.grid(row=i, column=0, sticky="w", padx=(5,0), pady=2)

            remove_btn = ctk.CTkButton(self.selected_files_frame, text="X", width=25, height=25,
                                       fg_color="#ff4d4d", hover_color="#ff6666",
                                       command=lambda f=f: self.remove_file(f))
            remove_btn.grid(row=i, column=1, padx=(5,5))

    # ----------------------------- Eliminar archivo -----------------------------
    def remove_file(self, file_path):
        if file_path in self.selected_files:
            self.selected_files.remove(file_path)
            self.update_selected_files_label()

    # ----------------------------- Seleccionar carpeta de salida -----------------------------
    def select_output_folder(self):
        # Abrir en la carpeta del primer archivo seleccionado, si existe
        initial_dir = os.path.dirname(self.selected_files[0]) if self.selected_files else self.output_folder
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            self.output_folder = folder
            self.output_folder_label.configure(text=f"Carpeta de salida: {self.output_folder}")

    # ----------------------------- Procesar archivos -----------------------------
    def procesar_archivos(self):
        if not self.selected_files:
            return

        output_folder = self.output_folder
        processed_names = []

        for file_path in self.selected_files:
            try:
                raw_df = pd.read_excel(file_path, header=None)
                header_row = raw_df[raw_df.eq("Pes M").any(axis=1)].index[0]
                df = pd.read_excel(file_path, header=header_row)
                df = df.dropna(how="all").reset_index(drop=True)

                filename = os.path.basename(file_path)
                hop_number = re.search(r"HOP\s*([0-9]+)", filename, re.IGNORECASE)
                hop_number = hop_number.group(1) if hop_number else "UNKNOWN"

                # Convertir columna Tipus a formato MRAG
                if "Tipus" in df.columns:
                    df["Tipus"] = df["Tipus"].replace({"HG": "DWT", "EV": "GGWT"})

                # Convertir columnas numéricas
                for col in ["Pes M", "Pes", "Llarg", "Ample"]:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace(",", ".", regex=False)
                        df[col] = pd.to_numeric(df[col], errors="coerce")

                # ----------------- Formato sin pesos estimados -----------------
                if self.format_no_weights_var.get() == 1:
                    df['Pes individual'] = 1
                    for i in range(len(df) - 1):
                        pes_current = df.loc[i, "Pes M"]
                        pes_next = df.loc[i+1, "Pes M"]
                        if (pd.isna(pes_current) or pes_current == 0) and pd.notna(pes_next) and pes_next > 0:
                            df.loc[i, "Pes individual"] = 0
                            df.loc[i+1, "Pes individual"] = 0
                    df["Pes MRAG"] = df.apply(lambda row: row["Pes M"] if row["Pes individual"] == 1 else "NA", axis=1)
                    output_file = os.path.join(output_folder, f"Filtered_HOP{hop_number}.xlsx")
                    df.to_excel(output_file, index=False)
                    processed_names.append(f"Formato sin pesos estimados: {filename}")

                # ----------------- Calcular PesIndiv -----------------
                if self.calc_pesindiv_var.get() == 1:
                    df["Pes individual"] = 1
                    for i in range(len(df) - 1):
                        pes_current = df.loc[i, "Pes M"] if pd.notna(df.loc[i, "Pes M"]) else 0
                        pes_next = df.loc[i+1, "Pes M"] if pd.notna(df.loc[i+1, "Pes M"]) else 0
                        if pes_current == 0 and pes_next > 0:
                            df.loc[i, "Pes individual"] = 0
                            df.loc[i+1, "Pes individual"] = 0
                    df["Pes MRAG"] = df.apply(lambda row: row["Pes M"] if row["Pes individual"] == 1 and row["Pes M"] != 0 else "NA", axis=1)
                    output_file = os.path.join(output_folder, f"Individual_HOP{hop_number}.xlsx")
                    df.to_excel(output_file, index=False)
                    processed_names.append(f"PesIndiv calculado: {filename}")

            except Exception as e:
                processed_names.append(f"Error procesando {filename}: {e}")

        for widget in self.processed_files_frame.winfo_children():
            widget.destroy()

        for i, text in enumerate(processed_names):
            lbl = ctk.CTkLabel(self.processed_files_frame, text=text, anchor="w", justify="left")
            lbl.grid(row=i, column=0, sticky="w", padx=5, pady=2)

# -----------------------------
# Ejecutar app
# -----------------------------
if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()
