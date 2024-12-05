import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def load_data_from_excel(file_path):

    df = pd.read_excel(file_path, sheet_name="IRIS-TABELA", header=6)

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
        raise ValueError(
            f"As colunas {required_columns} não foram encontradas no arquivo."
        )

    rows_data = []
    thickness_data = []

    grouped = df.groupby("ROW")
    for row, group in grouped:
        num_tubes = len(group)
        thicknesses = group["THICKNESS (MM)"].values
        categories = group["LEGEND"].values

        rows_data.append((row, num_tubes, 0, categories))
        thickness_data.append(thicknesses)

    max_tubes = max([num_tubes for _, num_tubes, _, _ in rows_data])

    thickness_data = np.array(
        [
            np.pad(row, (0, max_tubes - len(row)), "constant", constant_values=np.nan)
            for row in thickness_data
        ]
    )

    return rows_data, thickness_data

def draw_heat_exchanger(
    rows_data,
    thickness_data,
    color_map,
    spacing_x,
    spacing_y,
    radius,
    nominal_thickness,
):
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.set_aspect("equal")

    max_num_tubes = max([num_tubes for _, num_tubes, _, _ in rows_data])

    total_width = (max_num_tubes - 1) * spacing_x + 2 * radius
    total_height = len(rows_data) * spacing_y + 2 * radius

    vertical_offsets = []

    for i, (row, num_tubes, _, categories) in enumerate(rows_data):
        if num_tubes == 0:
            continue

        vertical_offset_units = float(
            input(
                f"Digite o offset vertical para a linha {int(row)}: "
            )
        )
        horizontal_general_offset_units = float(
            input(
                f"Digite o offset horizontal geral para a linha {int(row)}: "
            )
        )
        horizontal_offset_units = float(
            input(
                f"Digite o offset horizontal entre os tubos da linha {int(row)}: "
            )
        )

        start_offset_tube = 0
        if horizontal_offset_units > 0:
            start_offset_tube = (
                int(
                    input(
                        f"A partir de qual tubo da linha {int(row)} o offset horizontal deve ser aplicado? "
                    )
                )
                - 1
            )

        vertical_offset = vertical_offset_units * radius
        horizontal_general_offset = horizontal_general_offset_units * radius   
        horizontal_offset = horizontal_offset_units * radius

        accumulated_vertical_offset = sum(vertical_offsets) + vertical_offset
        vertical_offsets.append(vertical_offset)

        y_position = (len(rows_data) - i - 1) * spacing_y - accumulated_vertical_offset

        for j in range(num_tubes):
            thickness = thickness_data[i, j]
            color = "white"

            if pd.notna(thickness):
                remaining_thickness_percentage = thickness / nominal_thickness

                for range_, col in color_map.items():
                    if (
                        isinstance(range_, tuple)
                        and range_[0] <= remaining_thickness_percentage <= range_[1]
                    ):
                        color = col
                        break

            if j <= start_offset_tube:
                x = j * spacing_x + horizontal_general_offset
            else:
                x = (
                (start_offset_tube + 1) * spacing_x
                + horizontal_general_offset
                + horizontal_offset
                + (j - start_offset_tube - 1) * spacing_x
    )

            circle = plt.Circle(
                (x, y_position), radius, color=color, ec="black", linewidth=0.5
            )
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
    ax.set_aspect("equal")
    ax.axis("off")
    plt.tight_layout()
    plt.show()

file_path = input(
    "Digite o caminho do arquivo Excel: "
)

rows_data, thickness_data = load_data_from_excel(file_path)


nominal_thickness = float(
    input("Digite a espessura nominal dos tubos do equipamento (em mm): ")
)

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
)
