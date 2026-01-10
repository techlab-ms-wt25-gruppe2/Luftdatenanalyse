import json

notebook_content = {
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# Luftqualitätsdaten: Import, Transformation & Excel-Export\n",
    "\n",
    "Dieses Notebook führt folgende Schritte aus:\n",
    "1. **Extraktion** von CSV-Daten aus verschachtelten ZIP-Archiven.\n",
    "2. **Bereinigung** und Kombination der Zeitstempel.\n",
    "3. **Pivotierung** in ein Analyse-freundliches Format (Wide-Format).\n",
    "4. **Export** der Ergebnisse in eine Excel-Datei."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "outputs": [],
   "source": [
    "import pandas as pd\n",
    "import os\n",
    "import zipfile\n",
    "\n",
    "data_dir = 'raw_data'\n",
    "li = []\n",
    "\n",
    "if not os.path.exists(data_dir):\n",
    "    print(f\"Ordner '{data_dir}' nicht gefunden. Bitte erstellen.\")\n",
    "else:\n",
    "    for zip_filename in os.listdir(data_dir):\n",
    "        if zip_filename.endswith(\".zip\"):\n",
    "            zip_path = os.path.join(data_dir, zip_filename)\n",
    "            with zipfile.ZipFile(zip_path, \"r\") as z:\n",
    "                for member_path in z.namelist():\n",
    "                    if member_path.endswith(\".csv\") and \"__MACOSX\" not in member_path:\n",
    "                        try:\n",
    "                            with z.open(member_path) as f:\n",
    "                                df_temp = pd.read_csv(\n",
    "                                    f, \n",
    "                                    sep=';', \n",
    "                                    decimal=',', \n",
    "                                    encoding='utf-8'\n",
    "                                )\n",
    "                                li.append(df_temp)\n",
    "                        except Exception as e:\n",
    "                            print(f\"Fehler bei {member_path}: {e}\")\n",
    "\n",
    "    if li:\n",
    "        df = pd.concat(li, axis=0, ignore_index=True).drop_duplicates()\n",
    "        print(f\"Erfolgreich geladen: {len(df)} Zeilen.\")\n",
    "    else:\n",
    "        print(\"Keine Daten gefunden.\")"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Daten transformieren"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "outputs": [],
   "source": [
    "if 'df' in locals() and not df.empty:\n",
    "    # Zeitstempel konvertieren\n",
    "    df['Timestamp'] = pd.to_datetime(df['Datum'] + ' ' + df['Uhrzeit'], dayfirst=True, errors='coerce')\n",
    "    \n",
    "    # Pivotieren (Schadstoffe in Spalten)\n",
    "    index_cols = ['Stationscode', 'Stationsname', 'Stationsumgebung', 'Art der Station', 'Timestamp', 'Einheit']\n",
    "    df_wide = df.pivot_table(\n",
    "        index=index_cols, \n",
    "        columns='Schadstoff', \n",
    "        values='Messwert',\n",
    "        aggfunc='first'\n",
    "    ).reset_index()\n",
    "    \n",
    "    df_wide.columns.name = None\n",
    "    display(df_wide.head())"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Export nach Excel\n",
    "Wir speichern die bereinigten Daten nun als `.xlsx` Datei ab."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "outputs": [],
   "source": [
    "output_file = \"Luftqualitaet_Zusammenfassung.xlsx\"\n",
    "\n",
    "try:\n",
    "    # index=False verhindert, dass die Zeilennummern von Pandas mitgespeichert werden\n",
    "    df_wide.to_excel(output_file, index=False, engine='openpyxl')\n",
    "    print(f\"✅ Datei erfolgreich gespeichert: {output_file}\")\n",
    "except Exception as e:\n",
    "    print(f\"❌ Fehler beim Speichern: {e}\")"
   ]
  }
 ],
 "metadata": {
  "kernelspec": { "display_name": "Python 3", "name": "python3" },
  "language_info": { "name": "python", "version": "3.8" }
 },
 "nbformat": 4, "nbformat_minor": 4
}

with open("luftdaten_mit_excel.ipynb", 'w', encoding='utf-8') as f:
    json.dump(notebook_content, f, indent=1)
print("Das Notebook 'luftdaten_mit_excel.ipynb' wurde erstellt.")