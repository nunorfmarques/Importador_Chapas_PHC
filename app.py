import streamlit as st
import pandas as pd
import io

# Configuração da página e estilo visual
st.set_page_config(page_title="Conversor Laser PHC", page_icon="⚙️")

# CSS para esconder menus do Streamlit e profissionalizar o aspeto
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp { background-color: #f8f9fa; }
    </style>
    """, unsafe_allow_html=True)

# DICIONÁRIO MESTRE (O teu banco de dados de referências)
DB_LASER = {
    "S235JR": {1.5: ("11041095001", "CHAPA LISA ESP. 1,50MM S235 JR"), 2.0: ("11041095002", "CHAPA LISA ESP. 2,00MM S235 JR"), 2.5: ("11041095003", "CHAPA LISA ESP. 2,50MM S235 JR"), 3.0: ("11041095004", "CHAPA LISA ESP. 3,00MM S235 JR"), 4.0: ("11041095005", "CHAPA LISA ESP. 4,00MM S235 JR"), 6.0: ("11041095007", "CHAPA LISA ESP. 6,00MM S235 JR"), 8.0: ("11041095008", "CHAPA LISA ESP. 8,00MM S235 JR"), 10.0: ("11041095009", "CHAPA LISA ESP. 10,00MM S235 JR")},
    "S275JR": {2.0: ("11041095019", "CHAPA LISA ESP. 2,00MM S275 JR"), 3.0: ("11041095021", "CHAPA LISA ESP. 3,00MM S275 JR"), 4.0: ("11041095023", "CHAPA LISA ESP. 5,00MM S275 JR"), 5.0: ("11041095022", "CHAPA LISA ESP. 4,00MM S275 JR"), 6.0: ("11041095024", "CHAPA LISA ESP. 6,00MM S275 JR"), 10.0: ("11041012010", "CHAPA LISA ESP. 10,00MM S275 JR")},
    "GALVANIZADO": {1.5: ("11041095039", "CHAPA GALVANIZADA ESP. 1,50MM"), 2.0: ("11041095040", "CHAPA GALVANIZADA ESP. 2,00MM"), 3.0: ("11041095041", "CHAPA GALVANIZADA ESP. 3,00MM")},
    "ZINCOR": {0.5: ("11041095035", "CHAPA ELETROZINCADA ESP. 0,50MM"), 1.5: ("11041095036", "CHAPA ELETROZINCADA ESP. 1,50MM"), 2.0: ("11041095037", "CHAPA ELETROZINCADA ESP. 2,00MM")}
}

st.title("🛡️ Portal de Conversão Laser")
st.info("Carregue o relatório PE original para gerar o ficheiro de importação PHC.")

arquivo = st.file_uploader("Arraste o ficheiro .xls aqui", type=["xls", "xlsx"])

if arquivo:
    df = pd.read_excel(arquivo, header=None)
    
    # Lógica de Paragem e Extração (Offsets V4.1)
    limite = df[df.apply(lambda r: r.astype(str).str.contains('TOTAIS DA CHAPA').any(), axis=1)].index.min()
    starts = df[df.apply(lambda r: r.astype(str).str.contains('DADOS DE PEÇA').any(), axis=1)].index.tolist()
    
    final_rows = []
    for s in starts:
        if pd.notna(limite) and s >= limite: break
        try:
            peca_ref = str(df.iloc[s+2, 13]).strip()
            material = str(df.iloc[s+6, 13]).upper()
            qtd = int(float(str(df.iloc[s+7, 37]).replace(',', '.')))
            esp = float(str(df.iloc[s+9, 40]).replace(',', '.'))
            peso_raw = df.iloc[s+11, 40]
            peso_phc = "{:.3f}".format(float(str(peso_raw).replace(',', '.'))).replace('.', ',')

            # Mapeamento
            grupo = "S275JR" if "275" in material else ("GALVANIZADO" if "GALV" in material else ("ZINCOR" if "ZINC" in material or "ELETRO" in material else "S235JR"))
            ref_phc, des_phc = DB_LASER.get(grupo, {}).get(esp, ("⚠️ NÃO MAPEADO", f"{grupo} {esp}mm"))
            
            final_rows.append([ref_phc, des_phc, peca_ref, qtd, "und", "0,000", peso_phc])
        except: continue

    if final_rows:
        df_final = pd.DataFrame(final_rows, columns=['Ref', 'Design', 'Peca', 'Qtt', 'Unidade', 'preco', 'peso'])
        st.success(f"Foram encontradas {len(df_final)} peças prontas para importação.")
        st.dataframe(df_final, use_container_width=True)
        
        # Conversão para Excel em memória
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="💾 Descarregar Ficheiro para o PHC",
            data=buffer.getvalue(),
            file_name="importar_phc.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
