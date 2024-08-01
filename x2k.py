import logging
import pandas as pd
import simplekml
import warnings
import os

LOG_LEVEL = os.getenv("LOG_LEVEL", default="INFO")
logging.basicConfig(
    level=LOG_LEVEL, format="%(asctime)s %(levelname)s %(filename)s: %(message)s"
)
logger = logging.getLogger("x2k")
info = logger.info
debug = logger.debug
warnings.simplefilter(action="ignore", category=FutureWarning)

colNames = [
    "ST Sostegno",
    "Apparato",
    "Denominazione Linea",
    "Tensione",
    "Accessibilità Sostegno",
    "Priorità",
    "GEOMETRIA_SOSTEGNO",
    "TIPO_STRUTTURA",
    "UI",
    "CODICE_LINEA_SAP",
    "DESCRIZIONE_LINEA",
    "PALIFICAZIONE",
    "ATT_COND",
    "TIPO_ARMAMENTO",
    "LONGITUDINE_SOST_FINE",
    "LATITUDINE_SOST_FINE",
    "ACCESSIBILITA",
    "ALTEZZA_UTILE",
    "CONDUTTORE",
]

colNamesForTooltip = [
    "Denominazione Linea",
    "ST Sostegno",
    "Apparato",
    # "Denominazione Linea",
    "Tensione",
    "Accessibilità Sostegno",
    "Priorità",
    "GEOMETRIA_SOSTEGNO",
    "TIPO_STRUTTURA",
    "UI",
    # "CODICE_LINEA_SAP",
    # "DESCRIZIONE_LINEA",
    "PALIFICAZIONE",
    # "ATT_COND",
    "TIPO_ARMAMENTO",
    # "LONGITUDINE_SOST_FINE",
    # "LATITUDINE_SOST_FINE",
    "ACCESSIBILITA",
    "ALTEZZA_UTILE",
    # "CONDUTTORE"
]

GroupByColumn = "UI"
PinName = "Denominazione Linea"
# PinName = "ST Sostegno"

DecisionColumn = "Apparato"

import tkinter as tk
from tkinter import filedialog, messagebox


def create_gui():
    def select_input_dir():
        dirname = filedialog.askdirectory()
        if dirname:
            input_dir_var.set(dirname)

    def select_output_dir():
        dirname = filedialog.askdirectory()
        if dirname:
            output_dir_var.set(dirname)

    def submit():
        input_dirname = input_dir_var.get()
        output_dirname = output_dir_var.get()
        selected_columns = [col for col, var in checkbox_vars.items() if var.get()]

        if not input_dirname or not output_dirname:
            messagebox.showwarning(
                "Input Error", "Please select both input and output directories."
            )
            return

        global colNames, colNamesForTooltip
        colNamesForTooltip = selected_columns

        debug("Input dir: %s", input_dirname)
        debug("Output dir: %s", output_dirname)
        for file in os.listdir(input_dirname):
            fname, ext = os.path.splitext(file)

            if not ext == ".xlsx":
                info("Skipping to process file due to unsupported extension %s", file)
                continue

            directory = os.path.join(output_dirname, fname)
            os.makedirs(directory, exist_ok=True)
            debug("Directory: " + directory)

            current_file = os.path.join(input_dirname, file)
            info("Processing file %s", current_file)

            # data = pd.read_excel(current_file)
            # colNames = data.columns
            # colNamesForTooltip = data.columns
            # colNamesForTooltip = selected_columns
            info(colNames)
            info(colNamesForTooltip)
            data = pd.read_excel(current_file)

            # df = pd.DataFrame(data)
            df = pd.DataFrame(data, columns=colNames)
            df.fillna("NA", inplace=True)

            genera_kml(df, directory)

        messagebox.showinfo("Success", "Processing completed successfully!")

    root = tk.Tk()
    root.title("KML Generator")

    default_dirname = os.path.realpath(".")
    default_input_dir = os.path.join(default_dirname, "input")
    default_output_dir = os.path.join(default_dirname, "extract")

    input_dir_var = tk.StringVar(value=default_input_dir)
    output_dir_var = tk.StringVar(value=default_output_dir)

    tk.Label(root, text="Input Directory:").grid(row=0, column=0, sticky=tk.W)
    tk.Entry(root, textvariable=input_dir_var, width=50).grid(row=0, column=1)
    tk.Button(root, text="Browse", command=select_input_dir).grid(row=0, column=2)

    tk.Label(root, text="Output Directory:").grid(row=1, column=0, sticky=tk.W)
    tk.Entry(root, textvariable=output_dir_var, width=50).grid(row=1, column=1)
    tk.Button(root, text="Browse", command=select_output_dir).grid(row=1, column=2)

    tk.Label(root, text="Select Columns for Tooltip:").grid(
        row=2, column=0, sticky=tk.W
    )

    checkbox_vars = {
        col: tk.BooleanVar(value=(col in colNamesForTooltip)) for col in colNames
    }

    for i, col in enumerate(colNames):
        tk.Checkbutton(root, text=col, variable=checkbox_vars[col]).grid(
            row=3 + i // 3, column=i % 3, sticky=tk.W
        )

    tk.Button(root, text="Submit", command=submit).grid(
        row=3 + len(colNames) // 3, column=1
    )

    root.mainloop()


def makeTooltipRow(columnName, row):
    label = columnName.replace("_", " ").title() + ": "
    value = str(row[columnName])
    return "<hr><b>" + label + "</b> " + value


def genera_kml(df, path):
    linee = df[GroupByColumn].unique()
    diz = {}
    for linea in linee:
        diz[linea] = df.loc[df[GroupByColumn] == linea]
        kml = simplekml.Kml()
        for index, row in diz[linea].iterrows():
            point = kml.newpoint(
                name=row[PinName],
                coords=[(row["LONGITUDINE_SOST_FINE"], row["LATITUDINE_SOST_FINE"])],
            )

            # point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png'
            # point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png'
            # point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png'

            cell = str(row[DecisionColumn]).lower().strip()

            if "master" in cell:
                point.style.iconstyle.icon.href = (
                    "http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png"
                )
            elif "slave" in cell:
                point.style.iconstyle.icon.href = (
                    "http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png"
                )
            elif "nok" in cell:
                point.style.iconstyle.icon.href = (
                    "http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png"
                )
            elif "na" in cell:
                point.style.iconstyle.icon.href = (
                    "http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png"
                )
            else:
                print(cell)

            descrizione = ""

            for columnName in colNamesForTooltip:
                descrizione += makeTooltipRow(columnName, row)

            # descrizione += "<hr><b>Tipo armamento:</b> "+str(row["TIPO_ARMAMENTO"])
            # descrizione += "<hr><b>Accessibilità sostegni:</b> "+str(row["ACCESSIBILITA"])
            # descrizione += "<hr><b>Palificazione:</b> "+str(row["PALIFICAZIONE"])

            point.description = descrizione
        screen = kml.newscreenoverlay(name="Legenda")
        screen.icon.href = "legenda.png"
        screen.overlayxy = simplekml.OverlayXY(
            x=0, y=0, xunits=simplekml.Units.fraction, yunits=simplekml.Units.fraction
        )
        screen.screenxy = simplekml.ScreenXY(
            x=0, y=0, xunits=simplekml.Units.fraction, yunits=simplekml.Units.fraction
        )
        kml.save(str(path) + "/" + str(linea) + ".kml")

    df.to_excel(str(path) + "/" + "df.xlsx", index=False)


def main():
    info("Ciao!")
    create_gui()

    # warnings.simplefilter(action="ignore", category=FutureWarning)
    # dirname = os.path.realpath(".")
    # global colNames, colNamesForTooltip
    # output_dirname = os.path.join(dirname, "extract")
    # input_dirname = os.path.join(dirname, "input")
    #
    # if not os.path.exists(input_dirname):
    #     info("No file to process. Ciao ciao!")
    #     exit()
    #
    # for file in os.listdir(input_dirname):
    #     current_file = os.path.join(input_dirname, file)
    #     info("Processing file %s", current_file)
    #
    #     data = pd.read_excel(current_file)
    #     colNames = data.columns
    #     colNamesForTooltip = data.columns
    #     # df = pd.DataFrame(data, columns=colNames)
    #     df = pd.DataFrame(data)
    #     df.fillna("NA", inplace=True)
    #
    #     fname = file.split(os.path.extsep)[0]
    #     directory = os.path.join(output_dirname, fname)
    #
    #     os.makedirs(directory, exist_ok=True)
    #     debug("Directory: " + directory)
    #     genera_kml(df, directory)


if __name__ == "__main__":
    main()
