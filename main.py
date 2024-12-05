import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, Frame

def load_data_from_excel(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name="IRIS-TABELA", header=6, engine="openpyxl")
    except ValueError as e:
        raise ValueError("Erro ao abrir o arquivo Excel. Verifique se ele é compatível.") from e

    df.columns = df.columns.str.strip().str.upper()
    column_mapping = {
        "FILEIRA": "ROW",
        "TUBO": "TUBE",
        "ESPESSURA (MM)": "THICKNESS (MM)",
        "LEGENDA": "LEGEND",
    }
    df = df.rename(columns={col: column_mapping.get(col, col) for col in df.columns})
    required_columns = ["ROW", "TUBE", "THICKNESS (MM)", "LEGEND"]

    if not all(col in df.columns for col in required_columns):
        raise ValueError(f"As colunas {required_columns} não foram encontradas no arquivo.")

    rows_data, thickness_data = [], []
    grouped = df.groupby("ROW")
    for row, group in grouped:
        num_tubes = len(group)
        thicknesses = group["THICKNESS (MM)"].values
        rows_data.append((row, num_tubes))
        thickness_data.append(thicknesses)

    max_tubes = max([num_tubes for _, num_tubes in rows_data])
    thickness_data = np.array([
        np.pad(row, (0, max_tubes - len(row)), "constant", constant_values=np.nan)
        for row in thickness_data
    ])
    return rows_data, thickness_data

def draw_heat_exchanger(rows_data, thickness_data, color_map, spacing_x, spacing_y, radius, nominal_thickness, offsets):
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.set_aspect("equal")

    max_num_tubes = max([num_tubes for _, num_tubes in rows_data])
    total_width = (max_num_tubes - 1) * spacing_x + 2 * radius
    total_height = len(rows_data) * spacing_y + 2 * radius

    accumulated_vertical_offset = 0

    for i, (row, num_tubes) in enumerate(rows_data):
        if num_tubes == 0:
            continue

        offset_vertical = offsets[row]["offset_vertical"] * radius
        offset_horizontal_general = offsets[row]["offset_horizontal_general"] * radius
        offset_horizontal = offsets[row]["offset_horizontal"] * radius
        start_position_offset_tube = offsets[row]["start_position_offset_tube"] - 1

        accumulated_vertical_offset += offset_vertical
        y_position = (len(rows_data) - i - 1) * spacing_y - accumulated_vertical_offset

        for j in range(num_tubes):
            thickness = thickness_data[i, j]
            color = "white"

            if pd.notna(thickness):
                remaining_thickness_percentage = thickness / nominal_thickness
                for range_, col in color_map.items():
                    if isinstance(range_, tuple) and range_[0] <= remaining_thickness_percentage <= range_[1]:
                        color = col
                        break

            if j <= start_position_offset_tube:
                x = j * spacing_x + offset_horizontal_general
            else:
                x = (
                    (start_position_offset_tube + 1) * spacing_x
                    + offset_horizontal_general
                    + offset_horizontal
                    + (j - start_position_offset_tube - 1) * spacing_x
                )

            circle = plt.Circle((x, y_position), radius, color=color, ec="black", linewidth=0.5)
            ax.add_artist(circle)

            if j == 0:
                ax.text(
                    x - spacing_x - radius,
                    y_position,
                    f"{int(row)}",
                    ha="center",
                    va="center",
                    fontsize=9,
                )

    ax.set_xlim(-total_width / 2 - 2 * radius, total_width / 2 + 2 * radius)
    ax.set_ylim(-radius, total_height + 2 * radius)
    ax.axis("off")
    plt.tight_layout()
    plt.show()

def process_and_draw():
    try:
        file_path = file_path_entry.get()
        if not file_path:
            raise ValueError("Selecione um arquivo Excel válido.")

        nominal_thickness = float(nominal_thickness_entry.get())
        rows_data, thickness_data = load_data_from_excel(file_path)

        offsets = {}
        for row in rows_data:
            row_number = row[0]
            offset_vertical = float(offset_entries[row_number]["offset_vertical"].get())
            offset_horizontal_general = float(offset_entries[row_number]["offset_horizontal_general"].get())
            offset_horizontal = float(offset_entries[row_number]["offset_horizontal"].get())
            start_position_offset_tube = int(offset_entries[row_number]["start_position_offset_tube"].get())
            offsets[row_number] = {
                "offset_vertical": offset_vertical,
                "offset_horizontal_general": offset_horizontal_general,
                "offset_horizontal": offset_horizontal,
                "start_position_offset_tube": start_position_offset_tube,
            }

        color_map = {
            (0.8, 1.0): "blue",
            (0.6, 0.8): "green",
            (0.4, 0.6): "yellow",
            (0.2, 0.4): "orange",
            (0.0, 0.2): "red",
            "NI": "white",
            "FURO": "pink",
            "PLG": "black",
            "TA": "purple",
            "NDD": "cyan",
            "LIND": "brown",
            "OBT": "gray",
        }

        draw_heat_exchanger(
            rows_data=rows_data,
            thickness_data=thickness_data,
            color_map=color_map,
            spacing_x=0.9,
            spacing_y=0.9,
            radius=0.3,
            nominal_thickness=nominal_thickness,
            offsets=offsets,
        )
    except Exception as e:
        messagebox.showerror("Erro", str(e))

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xls *.xlsx *.xlsm *.xlsb")],
    )
    if file_path:
        file_path_entry.delete(0, "end")
        file_path_entry.insert(0, file_path)
        populate_offsets(file_path)  # Carregar os offsets após o arquivo ser selecionado

def populate_offsets(file_path):
    try:
        rows_data, _ = load_data_from_excel(file_path)
        
        for widget in offset_frame.winfo_children():
            widget.destroy()

        global offset_entries
        offset_entries = {}

        # Adiciona labels para os campos de input
        Label(offset_frame, text="Offset Vertical").grid(row=0, column=1, padx=5, pady=5)
        Label(offset_frame, text="Offset Horizontal Geral").grid(row=0, column=2, padx=5, pady=5)
        Label(offset_frame, text="Offset Horizontal Entre Tubos (mm)").grid(row=0, column=3, padx=5, pady=5)
        Label(offset_frame, text="Offset Tubo Inicial").grid(row=0, column=4, padx=5, pady=5)

        for i, (row, _) in enumerate(rows_data):
            row_number = int(row)

            # Coloca o número da linha à esquerda de cada caixa de input
            Label(offset_frame, text=f"Linha {row_number}").grid(row=i+1, column=0, padx=5, pady=5, sticky="e")

            # Entradas para os valores de offset
            offset_entries[row_number] = {
                "offset_vertical": Entry(offset_frame, width=10),
                "offset_horizontal_general": Entry(offset_frame, width=10),
                "offset_horizontal": Entry(offset_frame, width=10),
                "start_position_offset_tube": Entry(offset_frame, width=10),
            }

            # As entradas devem ser organizadas em colunas corretas
            offset_entries[row_number]["offset_vertical"].grid(row=i+1, column=1, padx=2, pady=2)
            offset_entries[row_number]["offset_horizontal_general"].grid(row=i+1, column=2, padx=2, pady=2)
            offset_entries[row_number]["offset_horizontal"].grid(row=i+1, column=3, padx=2, pady=2)
            offset_entries[row_number]["start_position_offset_tube"].grid(row=i+1, column=4, padx=2, pady=2)

    except Exception as e:
        messagebox.showerror("Erro", str(e))

root = Tk()
root.title("Heat Exchanger Visualizer")

Label(root, text="Arquivo Excel:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
file_path_entry = Entry(root, width=40)
file_path_entry.grid(row=0, column=1, padx=5, pady=5)

Button(root, text="Buscar Arquivo", command=browse_file).grid(row=0, column=2, padx=5, pady=5)

Label(root, text="Espessura Nominal (mm):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
nominal_thickness_entry = Entry(root, width=10)
nominal_thickness_entry.grid(row=1, column=1, padx=5, pady=5)

Button(root, text="Gerar Croqui", command=process_and_draw).grid(row=2, column=0, columnspan=3, pady=10)

offset_frame = Frame(root)
offset_frame.grid(row=3, column=0, columnspan=3, pady=10)

root.mainloop()
