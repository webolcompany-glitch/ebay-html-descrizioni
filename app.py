import streamlit as st
import pandas as pd
from io import BytesIO


st.set_page_config(
    page_title="eBay HTML Generator",
    layout="wide"
)


# =============================
# FUNZIONI GENERAZIONE
# =============================


def crea_lista(testo):

    if pd.isna(testo) or testo == "":
        return ""

    elementi = str(testo).split(";")

    html = ""

    for elemento in elementi:
        elemento = elemento.strip()

        if elemento:
            html += f"<li>{elemento}</li>\n"

    return html



def crea_tabella(testo):

    if pd.isna(testo) or testo == "":
        return ""

    righe = str(testo).split(";")

    html = ""

    for riga in righe:

        if ":" in riga:

            campo, valore = riga.split(":", 1)

            html += f"""
<tr>
<td>{campo.strip()}</td>
<td>{valore.strip()}</td>
</tr>
"""

    return html



def crea_descrizione(testo):

    if pd.isna(testo) or testo == "":
        return ""

    paragrafi = str(testo).split("\n")

    html = ""

    for p in paragrafi:

        if p.strip():

            html += f"""
<p>
{p.strip()}
</p>
"""

    return html



def crea_galleria(row):

    html = ""

    # FOTO4 - FOTO9

    for i in range(4, 10):

        colonna = f"FOTO{i}"

        if colonna in row:

            foto = row[colonna]

            if pd.notna(foto) and str(foto).strip() != "":

                html += f"""
<img src="{foto}" alt="Immagine prodotto">
"""


    return html



def genera_html(row, template):


    html = template



    sostituzioni = {


        "{{TITOLO}}":
        row.get("TITOLO",""),


        "{{SOTTOTITOLO}}":
        row.get("SOTTOTITOLO",""),


        "{{DESCRIZIONE}}":
        crea_descrizione(
            row.get("DESCRIZIONE","")
        ),


        "{{CARATTERISTICHE}}":
        crea_lista(
            row.get("CARATTERISTICHE","")
        ),


        "{{DATI_TECNICI}}":
        crea_tabella(
            row.get("DATI_TECNICI","")
        ),


        "{{CONTENUTO}}":
        crea_lista(
            row.get("CONTENUTO","")
        ),


        "{{NOTA}}":
        row.get("NOTA",""),



        "{{FOTO1}}":
        row.get("FOTO1",""),


        "{{FOTO2}}":
        row.get("FOTO2",""),


        "{{FOTO3}}":
        row.get("FOTO3",""),



        "{{GALLERIA}}":
        crea_galleria(row)

    }



    # sostituzione variabili

    for chiave, valore in sostituzioni.items():


        if pd.isna(valore):

            valore = ""


        html = html.replace(
            chiave,
            str(valore)
        )


    return html





# =============================
# INTERFACCIA STREAMLIT
# =============================


st.title("🟧 eBay HTML Generator")

st.write(
    "Generatore automatico descrizioni HTML eBay da Excel"
)



excel_file = st.file_uploader(
    "Carica file Excel prodotti",
    type=["xlsx"]
)



template_file = st.file_uploader(
    "Carica template HTML",
    type=["html"]
)




if excel_file and template_file:


    df = pd.read_excel(
        excel_file
    )


    template = template_file.read().decode(
        "utf-8"
    )


    st.subheader(
        "Anteprima dati caricati"
    )


    st.dataframe(
        df.head()
    )



    if st.button(
        "🚀 GENERA EXCEL CON HTML"
    ):



        html_finali = []



        barra = st.progress(0)



        totale = len(df)



        for indice, (_, row) in enumerate(df.iterrows()):


            html = genera_html(
                row,
                template
            )


            html_finali.append(
                html
            )


            barra.progress(
                (indice + 1) / totale
            )



        df["HTML_COMPLETO"] = html_finali



        output = BytesIO()



        df.to_excel(
            output,
            index=False,
            engine="openpyxl"
        )



        output.seek(0)



        st.success(
            "File Excel generato correttamente!"
        )



        st.download_button(

            label="⬇ SCARICA EXCEL EBAY",

            data=output,

            file_name="inserzioni_ebay_html.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )
