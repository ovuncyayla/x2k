import logging
import pandas as pd
import simplekml
import warnings
import os

LOG_LEVEL=os.getenv("LOG_LEVEL", default="INFO")
logging.basicConfig(level=LOG_LEVEL, format="%(asctime)s %(levelname)s %(filename)s: %(message)s")
logger = logging.getLogger("x2k")
info = logger.info
debug = logger.debug

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
    warnings.simplefilter(action="ignore", category=FutureWarning)
    dirname = os.path.realpath(".")
    output_dirname = os.path.join(dirname, "extract")
    input_dirname = os.path.join(dirname, "input")

    if not os.path.exists(input_dirname):
        info("No file to process. Ciao ciao!")
        exit()

    for file in os.listdir(input_dirname):
        current_file = os.path.join(input_dirname, file)
        info("Processing file %s", current_file)

        data = pd.read_excel(current_file)
        df = pd.DataFrame(data, columns=colNames)
        df.fillna("NA", inplace=True)

        fname = file.split(os.path.extsep)[0]
        directory = os.path.join(output_dirname, fname)

        os.makedirs(directory, exist_ok=True)
        debug("Directory: " + directory)
        genera_kml(df, directory)

if __name__ == "__main__":
    main()
