import streamlit as st
import pandas as pd
import io

# Configuração da Página
st.set_page_config(page_title="Conversor Laser PHC", page_icon="🚀", layout="centered")

# Estilo para esconder o menu do Streamlit
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# DICIONÁRIO MESTRE
DB_LASER = {
    "S235JR": {1.5: ("11041095001", "CHAPA LAMINADA A QUENTE LISA ESP. 1,50MM S235 JR EN 10025"), 2.0: ("11041095002", "CHAPA LAMINADA A QUENTE LISA ESP. 2,00MM S235 JR EN 10025"), 2.5: ("11041095003", "CHAPA LAMINADA A QUENTE LISA ESP. 2,50MM S235 JR EN 10025"), 3.0: ("11041095004", "CHAPA LAMINADA A QUENTE LISA ESP. 3,00MM S235 JR EN 10025"), 4.0: ("11041095005", "CHAPA LAMINADA A QUENTE LISA ESP. 4,00MM S235 JR EN 10025"), 5.0: ("11041095006", "CHAPA LAMINADA A QUENTE LISA ESP. 5,00MM S235 JR EN 10035"), 6.0: ("11041095007", "CHAPA LAMINADA A QUENTE LISA ESP. 6,00MM S235 JR EN 10025"), 8.0: ("11041095008", "CHAPA LAMINADA A QUENTE LISA ESP. 8,00MM S235 JR EN 10025"), 10.0: ("11041095009", "CHAPA LAMINADA A QUENTE LISA ESP. 10,00MM S235 JR EN 10025")},
    "S275JR": {2.0: ("11041095019", "CHAPA LAMINADA A QUENTE LISA ESP. 2,00MM S275 JR EN 10025"), 2.5: ("11041095020", "CHAPA LAMINADA A QUENTE LISA ESP. 2,50MM S275 JR EN 10025"), 3.0: ("11041095021", "CHAPA LAMINADA A QUENTE LISA ESP. 3,00MM S275 JR EN 10025"), 4.0: ("11041095022", "CHAPA LAMINADA A QUENTE LISA ESP. 4,00MM S275 JR EN 10025"), 6.0: ("11041095024", "CHAPA LAMINADA A QUENTE LISA ESP. 6,00MM S275 JR EN 10025"), 10.0: ("11041012010", "CHAPA LAMINADA A QUENTE LISA ESP. 10,00MM S275 JR EN 10025")},
    "GALVANIZADO": {1.5: ("11041095039", "CHAPA GALVANIZADA LISA ESP. 1,50MM DX51D EN 10327"), 2.0: ("11041095040", "CHAPA GALVANIZADA LISA ESP. 2,00MM DX51D EN 10327"), 3.0: ("11041095041", "CHAPA GALVANIZADA LISA ESP. 3,00MM DX51D EN 10327")},
    "ZINCOR": {0.5: ("11041095035", "CHAPA ELETROZINCADA LISA ESP. 0,50MM DC01 + ZE 25/25 EN 101"), 1.5: ("11041095036", "CHAPA ELETROZINCADA LISA ESP. 1,50MM DC01 + ZE 25/25 EN 101"), 2.0: ("11041095037", "CHAPA ELETROZINCADA LISA ESP. 2,00MM DC01 + ZE 25/25 EN 101")}
}

st.title("📂 Conversor Industrial Laser")
st.subheader("Extração de dados para PHC")

uploaded_file = st.file_uploader("Selecione o relatório PE (.xls)", type=["xls", "xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    
    # Lógica de Paragem e Extração
    limite = df[df.apply(lambda r: r.astype(str).str.contains('TOTAIS DA CHAPA').any(), axis=1)].index.min()
    starts = df[df.apply(lambda r: r.astype(str).str.contains('DADOS DE PEÇA').any(), axis=1)].index.tolist()
    
    final_data = []
    for s in starts:
        if pd.notna(limite) and s >= limite: break
        try:
            peca_ref = str(df.iloc[s+2, 13]).strip()
            material = str(df.iloc[s+6, 13]).upper()
            qtd = int(float(str(df.iloc[s+7, 37]).replace(',', '.')))
            esp = float(str(df.iloc[s+9, 40]).replace(',', '.'))
            peso_raw = df.iloc[s+11, 40]
            peso_phc = "{:.3f}".format(float(str(peso_raw).replace(',', '.'))).replace('.', ',')

            grupo = "S275JR" if "275" in material else ("GALVANIZADO" if "GALV" in material else ("ZINCOR" if "ZINC" in material or "ELETRO" in material else "S235JR"))
            ref_phc, des_phc = DB_LASER.get(grupo, {}).get(esp, ("⚠️ NÃO MAPEADO", f"{grupo} {esp}mm"))
            
            final_data.append([ref_phc, des_phc, peca_ref, qtd, "und", "0,000", peso_phc])
        except: continue

    if final_data:
        df_import = pd.DataFrame(final_data, columns=['Ref', 'Design', 'Peca', 'Qtt', 'Unidade', 'preco', 'peso'])
        st.success(f"Apuradas {len(df_import)} peças.")
        st.dataframe(df_import)
        
        # Preparar Download
        towrite = io.BytesIO()
        df_import.to_excel(towrite, index=False, engine='xlsxwriter')
        towrite.seek(0)
        
        st.download_button(
            label="📩 Descarregar Ficheiro para PHC",
            data=towrite,
            file_name="importacao_phc.xlsx",
            mime="application/vnd.ms-excel"
        )
