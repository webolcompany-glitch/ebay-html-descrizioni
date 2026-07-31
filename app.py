import streamlit as st
import pandas as pd
from io import BytesIO
import re


st.set_page_config(
    page_title="eBay HTML Generator",
    layout="wide"
)
# DOWNLOAD FILE ESEMPIO

try:

    with open(
        "esempio_file_input.xlsx",
        "rb"
    ) as file:

        st.download_button(

            label="📥 Scarica Excel esempio",

            data=file,

            file_name="esempio_file_input.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

except FileNotFoundError:

    st.warning(
        "esempio_file_input.xlsx non trovato"
    )

# ==================================
# CARICAMENTO TEMPLATE FISSO
# ==================================

with open(
    "template.html",
    "r",
    encoding="utf-8"
) as file:

    template = file.read()



# ==================================
# FUNZIONI
# ==================================


def valore_cella(row, campo):

    if campo in row:

        valore = row[campo]

        if pd.notna(valore):
            return str(valore)

    return ""



def crea_lista(testo):

    if not testo:
        return ""

    html = ""

    elementi = str(testo).split(";")

    for elemento in elementi:

        elemento = elemento.strip()

        if elemento:

            html += f"""
<li>{elemento}</li>
"""

    return html



def crea_descrizione(testo):

    if not testo:
        return ""

    html = ""

    paragrafi = str(testo).split("\n")


    for p in paragrafi:

        if p.strip():

            html += f"""
<p>
{p.strip()}
</p>
"""


    return html



def crea_tabella(testo):

    if not testo:
        return ""

    html = ""

    righe = str(testo).split(";")

    for riga in righe:


        if ":" in riga:

            campo, valore = riga.split(":",1)


            html += f"""

<tr>
<td>{campo.strip()}</td>
<td>{valore.strip()}</td>
</tr>

"""


    return html



def crea_galleria(row):

    html = ""


    # cerca automaticamente tutte le colonne FOTO

    colonne_foto = []


    for colonna in row.index:


        if re.match(
            r"FOTO\d+",
            str(colonna)
        ):

            numero = int(
                re.findall(
                    r"\d+",
                    colonna
                )[0]
            )


            # dalla FOTO4 in poi

            if numero >= 4:

                colonne_foto.append(
                    colonna
                )



    # ordina FOTO4 FOTO5 FOTO6...

    colonne_foto.sort(
        key=lambda x:int(
            re.findall(
                r"\d+",
                x
            )[0]
        )
    )



    for foto_colonna in colonne_foto:


        foto = valore_cella(
            row,
            foto_colonna
        )


        if foto:


            html += f"""

<img src="{foto}" alt="Immagine prodotto">

"""



    return html





def genera_html(row):


    html = template



    sostituzioni = {


        "{{TITOLO}}":
        valore_cella(row,"TITOLO"),


        "{{SOTTOTITOLO}}":
        valore_cella(row,"SOTTOTITOLO"),



        "{{DESCRIZIONE}}":
        crea_descrizione(
            valore_cella(
                row,
                "DESCRIZIONE"
            )
        ),



        "{{CARATTERISTICHE}}":
        crea_lista(
            valore_cella(
                row,
                "CARATTERISTICHE"
            )
        ),



        "{{DATI_TECNICI}}":
        crea_tabella(
            valore_cella(
                row,
                "DATI_TECNICI"
            )
        ),



        "{{CONTENUTO}}":
        crea_lista(
            valore_cella(
                row,
                "CONTENUTO"
            )
        ),



        "{{NOTA}}":
        valore_cella(
            row,
            "NOTA"
        ),



        "{{FOTO1}}":
        valore_cella(
            row,
            "FOTO1"
        ),



        "{{FOTO2}}":
        valore_cella(
            row,
            "FOTO2"
        ),



        "{{FOTO3}}":
        valore_cella(
            row,
            "FOTO3"
        ),



        "{{GALLERIA}}":
        crea_galleria(row)

    }



    for chiave,valore in sostituzioni.items():


        html = html.replace(
            chiave,
            valore
        )


    return html





# ==================================
# INTERFACCIA
# ==================================


st.title(
    "🟧 Generatore HTML eBay"
)



excel = st.file_uploader(
    "Carica Excel prodotti",
    type=["xlsx"]
)



if excel:


    df = pd.read_excel(
        excel
    )


    st.subheader(
        "Anteprima prodotti"
    )


    st.dataframe(
        df.head()
    )



    if st.button(
        "GENERARE HTML"
    ):


        risultati = []


        barra = st.progress(0)



        totale = len(df)



        for indice,(_,row) in enumerate(
            df.iterrows()
        ):


            html = genera_html(
                row
            )


            risultati.append(
                html
            )


            barra.progress(
                (indice+1)/totale
            )



        df["HTML_COMPLETO"] = risultati



        file_output = BytesIO()



        df.to_excel(
            file_output,
            index=False,
            engine="openpyxl"
        )



        file_output.seek(0)



        st.success(
            "File creato!"
        )



        st.download_button(

            "⬇ Scarica Excel finale",

            data=file_output,

            file_name="ebay_html_generato.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )
