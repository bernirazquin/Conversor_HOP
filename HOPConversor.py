import customtkinter as ctk
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
import pythoncom
from win32com.shell import shell

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
    "ES": {"title": "Procesador Excel HOP", "select_button": "Seleccionar archivos",
           "step1": "Step 1: Calcular PesIndiv y normalizar Tipus",
           "step2": "Step 2: Formato Word Sampling",
           "step3": "Step 3: Formato Access Sampling",
           "choose_folder": "Elegir carpeta de salida",
           "process": "Procesar archivos",
           "files_selected": "Archivos seleccionados:",
           "files_processed": "Archivos procesados:"},
    "EN": {"title": "HOP Excel Processor", "select_button": "Select files",
           "step1": "Step 1: Calculate PesIndiv and normalize Tipus",
           "step2": "Step 2: Word Sampling Format",
           "step3": "Step 3: Access Sampling Format",
           "choose_folder": "Choose output folder",
           "process": "Process files",
           "files_selected": "Selected files:",
           "files_processed": "Processed files:"},
    "HR": {"title": "HOP Excel Procesor", "select_button": "Odaberi datoteke",
           "step1": "Korak 1: Izračunaj PesIndiv i normaliziraj Tipus",
           "step2": "Korak 2: Word Sampling Format",
           "step3": "Korak 3: Access Sampling Format",
           "choose_folder": "Odaberi izlaznu mapu",
           "process": "Obradi datoteke",
           "files_selected": "Odabrane datoteke:",
           "files_processed": "Obrađene datoteke:"},
    "PT": {"title": "Processador Excel HOP", "select_button": "Selecionar arquivos",
           "step1": "Step 1: Calcular PesIndiv e normalizar Tipus",
           "step2": "Step 2: Formato Word Sampling",
           "step3": "Step 3: Formato Access Sampling",
           "choose_folder": "Escolher pasta de saída",
           "process": "Processar arquivos",
           "files_selected": "Arquivos selecionados:",
           "files_processed": "Arquivos processados:"}
}

# -----------------------------
# Clase principal de la app
# -----------------------------
class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("900x650")
        self.title(LANG_DICT["ES"]["title"])
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.current_lang = "ES"
        self.files = []
        self.output_folder = os.getcwd()

        # -----------------------------
        # Contenedor central
        # -----------------------------
        self.center_frame = ctk.CTkFrame(self, fg_color="#f0f0f0", corner_radius=20)
        self.center_frame.pack(pady=10, padx=30, fill="both", expand=True)

        # -----------------------------
        # Menú de idioma
        # -----------------------------
        self.lang_option = ctk.CTkOptionMenu(self, values=["ES","EN","HR","PT"],
                                             command=self.change_language, width=120)
        self.lang_option.place(relx=0.95, y=10, anchor="ne")
        self.lang_option.tkraise()

        # -----------------------------
        # Botón seleccionar archivos
        # -----------------------------
        self.select_button = ctk.CTkButton(
            self.center_frame,
            text=LANG_DICT[self.current_lang]["select_button"],
            command=self.select_files,
            corner_radius=15,
            width=220
        )
        self.select_button.pack(pady=(20,10))

        # -----------------------------
        # Label archivos seleccionados
        # -----------------------------
        self.selected_files_label = ctk.CTkLabel(
            self.center_frame,
            text="",
            fg_color="#e6e6e6",
            text_color="#003366",
            corner_radius=15,
            height=100,
            anchor="center",
            justify="center"
        )
        self.selected_files_label.pack(fill="x", padx=20, pady=5)

        # -----------------------------
        # Checkboxes para steps
        # -----------------------------
        self.step1_var = ctk.IntVar()
        self.step2_var = ctk.IntVar()
        self.step3_var = ctk.IntVar()
        self.step1 = ctk.CTkCheckBox(self.center_frame, text=LANG_DICT[self.current_lang]["step1"],
                                     variable=self.step1_var)
        self.step1.pack(pady=5)
        self.step2 = ctk.CTkCheckBox(self.center_frame, text=LANG_DICT[self.current_lang]["step2"],
                                     variable=self.step2_var)
        self.step2.pack(pady=5)
        self.step3 = ctk.CTkCheckBox(self.center_frame, text=LANG_DICT[self.current_lang]["step3"],
                                     variable=self.step3_var)
        self.step3.pack(pady=5)

        # -----------------------------
        # Botón elegir carpeta de salida
        # -----------------------------
        self.folder_button = ctk.CTkButton(
            self.center_frame,
            text=LANG_DICT[self.current_lang]["choose_folder"],
            command=self.select_output_folder,
            corner_radius=15,
            width=220
        )
        self.folder_button.pack(pady=10)

        # -----------------------------
        # Botón procesar archivos
        # -----------------------------
        self.process_button = ctk.CTkButton(
            self.center_frame,
            text=LANG_DICT[self.current_lang]["process"],
            command=self.procesar_archivos,
            corner_radius=15,
            width=220
        )
        self.process_button.pack(pady=10)

        # -----------------------------
        # Label archivos procesados
        # -----------------------------
        self.processed_files_label = ctk.CTkLabel(
            self.center_frame,
            text="",
            fg_color="#e6e6e6",
            text_color="#003366",
            corner_radius=15,
            height=100,
            anchor="center",
            justify="center"
        )
        self.processed_files_label.pack(fill="x", padx=20, pady=10)

        # -----------------------------
        # Firma de la app
        # -----------------------------
        self.signature_label = ctk.CTkLabel(
            self,
            text="Creado por Bernardo R. para los observadores de MRAG en la granja de L´Ametlla de Mar",
            text_color="#666666",
            font=("Arial", 9)
        )
        self.signature_label.pack(side="bottom", pady=10)

    # -----------------------------
    # Cambiar idioma
    # -----------------------------
    def change_language(self, choice):
        self.current_lang = choice
        self.title(LANG_DICT[choice]["title"])
        self.select_button.configure(text=LANG_DICT[choice]["select_button"])
        self.step1.configure(text=LANG_DICT[choice]["step1"])
        self.step2.configure(text=LANG_DICT[choice]["step2"])
        self.step3.configure(text=LANG_DICT[choice]["step3"])
        self.folder_button.configure(text=LANG_DICT[choice]["choose_folder"])
        self.process_button.configure(text=LANG_DICT[choice]["process"])

    # -----------------------------
    # Seleccionar archivos
    # -----------------------------
    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        for path in paths:
            real_path = resolve_lnk(path)
            if real_path not in self.files:
                self.files.append(real_path)
        self.update_selected_files_label()

    # -----------------------------
    # Actualizar label archivos seleccionados
    # -----------------------------
    def update_selected_files_label(self):
        if self.files:
            text = f"{LANG_DICT[self.current_lang]['files_selected']}\n" + "\n".join([os.path.basename(f) for f in self.files])
        else:
            text = ""
        self.selected_files_label.configure(text=text)

    # -----------------------------
    # Seleccionar carpeta de salida
    # -----------------------------
    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder

    # -----------------------------
    # Función principal de procesamiento
    # -----------------------------
    def procesar_archivos(self):
        if not self.files:
            return

        # Carpeta única de salida
        output_folder = os.path.join(self.output_folder, "HOP_Procesados")
        os.makedirs(output_folder, exist_ok=True)

        processed_names = []

        for file_path in self.files:
            try:
                raw_df = pd.read_excel(file_path, header=None)
                header_row = raw_df[raw_df.eq("Pes M").any(axis=1)].index[0]
                df = pd.read_excel(file_path, header=header_row)
                df = df.dropna(how="all").reset_index(drop=True)

                filename = os.path.basename(file_path)
                hop_number = filename.split("_HOP")[-1].split(".")[0]

                # ----------------- STEP 1 -----------------
                if self.step1_var.get() == 1:
                    if "Tipus" in df.columns:
                        df["Tipus"] = df["Tipus"].replace({"HG": "DWT", "EV": "GGWT"})
                    col_index = df.columns.get_loc("Pes M")
                    df.insert(col_index + 1, "PesIndiv", None)
                    for i in range(len(df)):
                        pesM = df.loc[i, "Pes M"]
                        if pd.isna(pesM):
                            df.loc[i, "PesIndiv"] = "Null"
                            continue
                        if i > 0 and pd.isna(df.loc[i-1, "Pes M"]):
                            pes1 = df.loc[i-1, "Pes"]
                            pes2 = df.loc[i, "Pes"]
                            if pd.isna(pes1) or pd.isna(pes2) or pes1 < 0 or pes2 < 0:
                                df.loc[i-1, "PesIndiv"] = "Null"
                                df.loc[i, "PesIndiv"] = "Null"
                            else:
                                p1 = round(pes1 / (pes1 + pes2) * pesM)
                                p2 = int(pesM) - p1
                                df.loc[i-1, "PesIndiv"] = p1 if p1 >= 0 else "Null"
                                df.loc[i, "PesIndiv"] = p2 if p2 >= 0 else "Null"
                        else:
                            df.loc[i, "PesIndiv"] = int(round(pesM)) if pesM >= 0 else "Null"
                    out_file_step1 = os.path.join(output_folder, f"Excel_Sacrifici_indiv_HOP{hop_number}.xlsx")
                    df.to_excel(out_file_step1, index=False)
                    processed_names.append(os.path.basename(out_file_step1))

                # ----------------- STEP 2 -----------------
                if self.step2_var.get() == 1:
                    df_step2 = df.copy()
                    tipus_clean = df_step2.get("Tipus", pd.Series()).astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "None": np.nan})
                    keep_mask = df_step2.get("Llarg", pd.Series()).notna() | tipus_clean.notna()
                    if not keep_mask.any():
                        processed_names.append(f"Step2 skipped: {filename}")
                        continue
                    df_step2 = df_step2.loc[keep_mask].reset_index(drop=True)
                    sampling_df = pd.DataFrame()
                    sampling_df["Fish ID"] = df_step2.get("Matrícula")
                    sampling_df["Length (cm)"] = df_step2.get("Llarg")
                    sampling_df["Length type (CFL/SFL)"] = "SFL"
                    sampling_df["Weight (Kg)"] = df_step2.get("Pes M")
                    sampling_df["PesIndiv"] = df_step2.get("PesIndiv")
                    sampling_df["Processed code"] = df_step2.get("Tipus")
                    sampling_df["Tag number/s"] = "NA"
                    sampling_df["Tag Type"] = "NA"
                    sampling_df["Condition around Tag"] = "NA"
                    out_file_step2 = os.path.join(output_folder, f"Sampling_format_HOP{hop_number}.xlsx")
                    sampling_df.to_excel(out_file_step2, index=False)
                    processed_names.append(os.path.basename(out_file_step2))

                # ----------------- STEP 3 -----------------
                if self.step3_var.get() == 1:
                    df_step3 = sampling_df if self.step2_var.get() == 1 else df.copy()
                    keep_mask = df_step3.get("Length (cm)", pd.Series()).notna()
                    if not keep_mask.any():
                        processed_names.append(f"Step3 skipped: {filename}")
                        continue
                    df_step3 = df_step3.loc[keep_mask].reset_index(drop=True)
                    access_df = pd.DataFrame()
                    access_df["Weight (kg)"] = df_step3.get("Weight (Kg)")
                    access_df["PesIndiv"] = df_step3.get("PesIndiv")
                    access_df["Processed code"] = df_step3.get("Processed code")
                    access_df["Length"] = df_step3.get("Length (cm)")
                    access_df["Length Type"] = df_step3.get("Length type (CFL/SFL)","SFL")
                    access_df["Tag number/s"] = "NA"
                    access_df["Tag Type"] = "NA"
                    access_df["Condition around Tag"] = "NA"
                    access_df["Biological Sample Collected"] = "NA"
                    access_df["Biological Sample Type"] = "NA"
                    access_df["Sample ID"] = "NA"
                    out_file_step3 = os.path.join(output_folder, f"sampling_access_HOP{hop_number}.xlsx")
                    access_df.to_excel(out_file_step3, index=False)
                    processed_names.append(os.path.basename(out_file_step3))

            except Exception as e:
                processed_names.append(f"Error: {filename}")
                continue

        if processed_names:
            text = f"{LANG_DICT[self.current_lang]['files_processed']}\n" + "\n".join(processed_names)
            text += f"\n\nGuardados en: {output_folder}"
            self.processed_files_label.configure(text=text)


# -----------------------------
# Ejecutar app
# -----------------------------
if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()
