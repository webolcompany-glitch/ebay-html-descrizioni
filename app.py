import streamlit as st
import pandas as pd
from io import BytesIO


st.set_page_config(
    page_title="eBay HTML Generator",
    layout="wide"
)



# -----------------------------
# FUNZIONI
# -----------------------------


def crea_lista(testo):

    if pd.isna(testo):
        return ""

    elementi = str(testo).split(";")

    risultato = ""

    for e in elementi:
        risultato += f"<li>{e.strip()}</li>\n"

    return risultato




def crea_tabella(testo):

    if pd.isna(testo):
        return ""

    righe = str(testo).split(";")

    html = ""

    for riga in righe:

        if ":" in riga:

            campo,valore = riga.split(":",1)

            html += f"""
<tr>
<td>{campo.strip()}</td>
<td>{valore.strip()}</td>
</tr>
"""

    return html





def crea_galleria(row):

    html=""

    for i in range(2,7):

        foto=row.get(f"FOTO{i}")

        if pd.notna(foto):

            html += f"""
<img src="{foto}" alt="Immagine prodotto">
"""

    return html





def genera_html(row,template):


    html=template



    sostituzioni={


    "{{TITOLO}}":
    row["TITOLO"],


    "{{SOTTOTITOLO}}":
    row["SOTTOTITOLO"],


    "{{DESCRIZIONE}}":
    row["DESCRIZIONE"],


    "{{CARATTERISTICHE}}":
    crea_lista(row["CARATTERISTICHE"]),


    "{{DATI_TECNICI}}":
    crea_tabella(row["DATI_TECNICI"]),


    "{{CONTENUTO}}":
    crea_lista(row["CONTENUTO"]),


    "{{NOTA}}":
    row["NOTA"],



    "{{FOTO1}}":
    row["FOTO1"],



    "{{FOTO2}}":
    row["FOTO2"],



    "{{FOTO3}}":
    row["FOTO3"],



    "{{FOTO4}}":
    row["FOTO4"],



    "{{FOTO5}}":
    row["FOTO5"],



    "{{FOTO6}}":
    row["FOTO6"],



    "{{GALLERIA}}":
    crea_galleria(row)

    }



    for chiave,valore in sostituzioni.items():

        if pd.isna(valore):
            valore=""

        html=html.replace(
            chiave,
            str(valore)
        )


    return html




# -----------------------------
# INTERFACCIA
# -----------------------------


st.title("🟧 eBay HTML Generator")

st.write(
"Genera automaticamente descrizioni HTML eBay dentro Excel"
)



file_excel = st.file_uploader(
    "Carica file Excel prodotti",
    type=["xlsx"]
)



file_template = st.file_uploader(
    "Carica template HTML",
    type=["html"]
)



if file_excel and file_template:


    df=pd.read_excel(file_excel)



    template=file_template.read().decode(
        "utf-8"
    )



    st.subheader(
        "Anteprima prodotti"
    )


    st.dataframe(
        df.head()
    )



    if st.button(
        "🚀 GENERA EXCEL HTML"
    ):


        lista_html=[]


        for _,row in df.iterrows():


            lista_html.append(
                genera_html(
                    row,
                    template
                )
            )



        df["HTML_COMPLETO"]=lista_html



        output=BytesIO()



        df.to_excel(
            output,
            index=False,
            engine="openpyxl"
        )



        output.seek(0)



        st.success(
            "File creato correttamente!"
        )



        st.download_button(

            label="⬇ Scarica Excel con HTML",

            data=output,

            file_name="prodotti_html_ebay.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )
