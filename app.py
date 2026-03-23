import streamlit as st
import pandas as pd
import io
import os

# Configuração da página
st.set_page_config(page_title="Importador Laser", page_icon="⚙")

# Esconder menus (Blindagem)
st.markdown("<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;}</style>", unsafe_allow_html=True)

# --- 1. CARREGAR DICIONÁRIO SHP (DO GITHUB) ---
@st.cache_data # Isto faz com que o ficheiro só seja lido uma vez, ficando rápido
def carregar_base_dados():
    caminho = "base_dados_shp.xls" # Nome exato do ficheiro no teu GitHub
    if os.path.exists(caminho):
        df_db = pd.read_excel(caminho, header=None)
        # Cria o mapeamento: Coluna 0 (SHP) -> Coluna 1 (Nome Real)
        return pd.Series(df_db.iloc[:, 1].values, index=df_db.iloc[:, 0].values).to_dict()
    return {}

dict_nomes = carregar_base_dados()

# DICIONÁRIO MESTRE (O teu banco de dados de referências)
DB_LASER = {
    "S235JR": {1.5: ("011041095001", "CHAPA LISA ESP. 1,50MM S235 JR"),
               2.0: ("011041095002", "CHAPA LISA ESP. 2,00MM S235 JR"),
               2.5: ("011041095003", "CHAPA LISA ESP. 2,50MM S235 JR"),
               3.0: ("011041095004", "CHAPA LISA ESP. 3,00MM S235 JR"),
               4.0: ("011041095005", "CHAPA LISA ESP. 4,00MM S235 JR"),
               6.0: ("011041095007", "CHAPA LISA ESP. 6,00MM S235 JR"),
               8.0: ("011041095008", "CHAPA LISA ESP. 8,00MM S235 JR"),
               10.0: ("011041095009", "CHAPA LISA ESP. 10,00MM S235 JR")},
    "S275JR": {
    2.0: ("011041095019", "CHAPA LISA ESP. 2,00MM S275 JR"),
    2.5: ("011041095020", "CHAPA LISA ESP. 2,50MM S275 JR"),
    3.0: ("011041095021", "CHAPA LISA ESP. 3,00MM S275 JR"),
    4.0: ("011041095022", "CHAPA LISA ESP. 4,00MM S275 JR"),
    5.0: ("011041095023", "CHAPA LISA ESP. 5,00MM S275 JR"),
    6.0: ("011041095024", "CHAPA LISA ESP. 6,00MM S275 JR"),
    8.0: ("011041095025", "CHAPA LISA ESP. 8,00MM S275 JR"),
    10.0: ("011041012010", "CHAPA LISA ESP. 10,00MM S275 JR"),
    12.0: ("011041095026", "CHAPA LISA ESP. 12,00MM S275 JR"),
    15.0: ("011041095027", "CHAPA LISA ESP. 15,00MM S275 JR"),
    16.0: ("011041095028", "CHAPA LISA ESP. 16,00MM S275 JR"),
    18.0: ("011041095029", "CHAPA LISA ESP. 18,00MM S275 JR"),
    20.0: ("011041095030", "CHAPA LISA ESP. 20,00MM S275 JR"),
    25.0: ("011041095031", "CHAPA LISA ESP. 25,00MM S275 JR"),
    30.0: ("011041095032", "CHAPA LISA ESP. 30,00MM S275 JR"),
    40.0: ("011041095033", "CHAPA LISA ESP. 40,00MM S275 JR"),
    50.0: ("011041095034", "CHAPA LISA ESP. 50,00MM S275 JR"),
},    
  
    "GALVANIZADO": {
                    1.5: ("011041095039", "CHAPA GALVANIZADA ESP. 1,50MM"),
                    2.0: ("011041095040", "CHAPA GALVANIZADA ESP. 2,00MM"),
                    3.0: ("011041095041", "CHAPA GALVANIZADA ESP. 3,00MM")
                   },
    
    "ZINCOR": {
                    0.5: ("011041095035", "CHAPA ELETROZINCADA ESP. 0,50MM"),
                    1.5: ("011041095036", "CHAPA ELETROZINCADA ESP. 1,50MM"),
                    2.0: ("011041095037", "CHAPA ELETROZINCADA ESP. 2,00MM"),
                    3.0: ("011041095038", "CHAPA ELETROZINCADA ESP. 3,00MM")
    }
}

st.title("🛡️ Portal de Importação Laser")
st.info("A base de dados de peças foi carregada automaticamente do sistema.")

# --- 2. UPLOAD DO RELATÓRIO DO LASER ---
arquivo_laser = st.file_uploader("Selecione o relatório PE (.xls)", type=["xls", "xlsx"])

if arquivo_laser:
    df = pd.read_excel(arquivo_laser, header=None)
    
    limite = df[df.apply(lambda r: r.astype(str).str.contains('TOTAIS DA CHAPA').any(), axis=1)].index.min()
    starts = df[df.apply(lambda r: r.astype(str).str.contains('DADOS DE PEÇA').any(), axis=1)].index.tolist()
    
    final_data = []
    for s in starts:
        if pd.notna(limite) and s >= limite: break
        try:
            shp_ref = str(df.iloc[s+2, 13]).strip()
            material = str(df.iloc[s+6, 13]).upper()
            qtd = int(float(str(df.iloc[s+7, 37]).replace(',', '.')))
            esp = float(str(df.iloc[s+9, 40]).replace(',', '.'))
            peso_raw = df.iloc[s+11, 40]
            peso_phc = "{:.3f}".format(float(str(peso_raw).replace(',', '.'))).replace('.', ',')

            # SUBSTITUIÇÃO AUTOMÁTICA (VLOOKUP INTERNO)
            nome_final = dict_nomes.get(shp_ref, shp_ref)

            # Lógica de Material
            if "275" in material: grupo = "S275JR"
            elif "GALV" in material: grupo = "GALVANIZADO"
            else: grupo = "S235JR"
            
            ref_phc, des_phc = DB_LASER.get(grupo, {}).get(esp, ("⚠️ NÃO MAP.", f"{grupo} {esp}mm"))
            
            final_data.append([ref_phc, des_phc, nome_final, qtd, "und", "0,000", peso_phc])
        except: continue

    if final_data:
        df_final = pd.DataFrame(final_data, columns=['Ref', 'Design', 'Peca', 'Qtt', 'Unidade', 'preco', 'peso'])
        st.success("✅ Processamento concluído!")
        st.dataframe(df_final, use_container_width=True)
        
        buffer = io.BytesIO()
        df_final.to_excel(buffer, index=False, engine='xlsxwriter')
        st.download_button("💾 Descarregar para PHC", buffer.getvalue(), "importacao_phc_revisto.xlsx")
