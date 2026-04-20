import streamlit as st
import pandas as pd
import io
import os
import zipfile

# Configuração da página
st.set_page_config(page_title="Importador Laser", page_icon="⚙")

# Esconder menus (Blindagem)
st.markdown("<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;}</style>", unsafe_allow_html=True)

# --- 1. CARREGAR DICIONÁRIO SHP (DO GITHUB) ---
@st.cache_data # Isto faz com que o ficheiro só seja lido uma vez, ficando rápido
def carregar_base_dados():
    caminho = "base_dados_shp.xlsx" # Nome exato do ficheiro no teu GitHub
    if os.path.exists(caminho):
        df_db = pd.read_excel(caminho, header=None)
        # Cria o mapeamento: Coluna 0 (SHP) -> Coluna 1 (Nome Real)
        return pd.Series(df_db.iloc[:, 1].values, index=df_db.iloc[:, 0].values).to_dict()
    return {}

dict_nomes = carregar_base_dados()

# DICIONÁRIO MESTRE (O teu banco de dados de referências)
DB_LASER = {
    "S235JR": {1.5: ("011041011011", "CHAPA LAMINADA A QUENTE LISA ESP. 1,50MM S235 JR EN 10025"),
               2.0: ("011041011012", "CHAPA LAMINADA A QUENTE LISA ESP. 2,00MM S235 JR EN 10025"),
               2.5: ("011041011013", "CHAPA LAMINADA A QUENTE LISA ESP. 2,50MM S235 JR EN 10025"),
               3.0: ("011041011014", "CHAPA LAMINADA A QUENTE LISA ESP. 3,00MM S235 JR EN 10025"),
               4.0: ("011041011015", "CHAPA LAMINADA A QUENTE LISA ESP. 4,00MM S235 JR EN 10025"),
               5.0: ("011041011016", "CHAPA LAMINADA A QUENTE LISA ESP. 5,00MM S235 JR EN 10025"),
               6.0: ("011041011017", "CHAPA LAMINADA A QUENTE LISA ESP. 6,00MM S235 JR EN 10025"),
               8.0: ("011041011018", "CHAPA LAMINADA A QUENTE LISA ESP. 8,00MM S235 JR EN 10025"),
               10.0: ("011041011019", "CHAPA LAMINADA A QUENTE LISA ESP. 10,00MM S235 JR EN 10025") 
              },
    "S275JR": {
                2.0: ("011041012012", "CHAPA LAMINADA A QUENTE LISA ESP. 2,00MM S275 JR EN 10025"),
                2.5: ("011041012013", "CHAPA LAMINADA A QUENTE LISA ESP. 2,50MM S275 JR EN 10025"),
                3.0: ("011041012014", "CHAPA LAMINADA A QUENTE LISA ESP. 3,00MM S275 JR EN 10025"),
                4.0: ("011041012015", "CHAPA LAMINADA A QUENTE LISA ESP. 4,00MM S275 JR EN 10025"),
                5.0: ("011041012016", "CHAPA LAMINADA A QUENTE LISA ESP. 5,00MM S275 JR EN 10025"),
                6.0: ("011041012017", "CHAPA LAMINADA A QUENTE LISA ESP. 6,00MM S275 JR EN 10025"),
                8.0: ("011041012018", "CHAPA LAMINADA A QUENTE LISA ESP. 8,00MM S275 JR EN 10025"),
                10.0: ("011041012019", "CHAPA LAMINADA A QUENTE LISA ESP. 10,00MM S275 JR EN 10025"),
                12.0: ("011041012020", "CHAPA LAMINADA A QUENTE LISA ESP. 12,00MM S275 JR EN 10025"),
                15.0: ("011041012021", "CHAPA LAMINADA A QUENTE LISA ESP. 15,00MM S275 JR EN 10025"),
                16.0: ("011041012022", "CHAPA LAMINADA A QUENTE LISA ESP. 16,00MM S275 JR EN 10025"),
                18.0: ("011041012023", "CHAPA LAMINADA A QUENTE LISA ESP. 18,00MM S275 JR EN 10025"),
                20.0: ("011041012024", "CHAPA LAMINADA A QUENTE LISA ESP. 20,00MM S275 JR EN 10025"),
                25.0: ("011041012025", "CHAPA LAMINADA A QUENTE LISA ESP. 25,00MM S275 JR EN 10025"),
                30.0: ("011041012026", "CHAPA LAMINADA A QUENTE LISA ESP. 30,00MM S275 JR EN 10025"),
                40.0: ("011041012027", "CHAPA LAMINADA A QUENTE LISA ESP. 40,00MM S275 JR EN 10025"),
                50.0: ("011041012028", "CHAPA LAMINADA A QUENTE LISA ESP. 50,00MM S275 JR EN 10025")
            },    
  
    "GALVANIZADO": {
                    1.5: ("011041031017", "CHAPA GALVANIZADA LISA ESP. 1,50MM DX51D EN 10327"),
                    2.0: ("011041031018", "CHAPA GALVANIZADA LISA ESP. 2,00MM DX51D EN 10327"),
                    2.5: ("011041031019", "CHAPA GALVANIZADA LISA ESP. 2,50MM DX51D EN 10327"),
                    3.0: ("011041031020", "CHAPA GALVANIZADA LISA ESP. 3,00MM DX51D EN 10327")
            },
    
    "ZINCOR": {
                    0.5: ("011041021011", "CHAPA ELETROZINCADA LISA ESP. 0,50MM DC01 + ZE 25/25 EN 10152"),
                    1.0: ("011041021013", "CHAPA ELETROZINCADA LISA ESP. 1,00MM DC01 + ZE 25/25 EN 10152"),
                    1.5: ("011041021016", "CHAPA ELETROZINCADA LISA ESP. 1,50MM DC01 + ZE 25/25 EN 10152"),
                    2.0: ("011041021017", "CHAPA ELETROZINCADA LISA ESP. 2,00MM DC01 + ZE 25/25 EN 10152"),
                    2.5: ("011041021018", "CHAPA ELETROZINCADA LISA ESP. 2,50MM DC01 + ZE 25/25 EN 10152"),   
                    3.0: ("011041031019", "CHAPA ELETROZINCADA LISA ESP. 3,00MM DC01 + ZE 25/25 EN 10152")
            },
    "INOX": {
                    1.5: ("011041012042", "CHAPA INOX AISI 304 2B PVC 1.50MM"),
                    2.0: ("011041091011", "CHAPA INOX AISI 304 2B PVC 2.00MM"),
                    3.0: ("011041091012", "CHAPA INOX AISI 304 2B PVC 3.00MM"),
                    4.0: ("011041091012", "CHAPA INOX AISI 304 2B 4.00MM"),
                    5.0: ("011041091012", "CHAPA INOX AISI 304 2B 5.00MM"),
                    6.0: ("011041091012", "CHAPA INOX AISI 304 2B 6.00MM"),
                    8.0: ("011041091012", "CHAPA INOX AISI 304 LQ 8.00MM"),
                    10.0: ("011041091012", "CHAPA INOX AISI 304 LQ 10.00MM")
            }
}

st.title("🛡️ Portal de Importação Laser")
st.info("A base de dados de peças foi carregada automaticamente do sistema.")

# --- 2. UPLOAD DO RELATÓRIO DO LASER ---
arquivos_laser = st.file_uploader("Selecione os relatório PE (.xls)", type=["xls", "xlsx"], accept_multiple_files=True)

if arquivos_laser:
    arquivos_processados = {}

    for arquivo_laser in arquivos_laser:
        df = pd.read_excel(arquivo_laser, header=None)
        
        limite = df[df.apply(lambda r: r.astype(str).str.contains('TOTAIS DA CHAPA').any(), axis=1)].index.min()
        starts = df[df.apply(lambda r: r.astype(str).str.contains('DADOS DE PEÇA').any(), axis=1)].index.tolist()
        
        final_data = []
        for s in starts:
            if pd.notna(limite) and s >= limite: break
            try:
                shp_ref = str(df.iloc[s+2, 13]).strip()
                nome_real = dict_nomes.get(shp_ref)
                if pd.isna(nome_real) or str(nome_real).strip() == "" or nome_real is None:
                    nome_final = shp_ref  # Mantém o SHP original para rastreabilidade
                else:
                    nome_final = str(nome_real).strip()
                material = str(df.iloc[s+6, 13]).upper()
                qtd = int(float(str(df.iloc[s+7, 37]).replace(',', '.')))
                esp = float(str(df.iloc[s+9, 40]).replace(',', '.'))
                peso_raw = df.iloc[s+11, 40]
                peso_phc = "{:.3f}".format(float(str(peso_raw).replace(',', '.'))).replace('.', ',')
    
                # Lógica de Material
                if "275" in material: grupo = "S275JR"
                elif "INOX" in material: grupo = "INOX"
                elif "GALV" in material: grupo = "GALVANIZADO"
                elif "ZINC" in material or "ELETRO" in material: grupo = "ZINCOR"
                else: grupo = "S235JR"
                
                ref_phc, des_phc = DB_LASER.get(grupo, {}).get(esp, ("⚠️ NÃO MAP.", f"{grupo} {esp}mm"))
                
                final_data.append([ref_phc, des_phc, nome_final, qtd, "und", "0,000", peso_phc])
            except: continue

        if final_data:
            df_final = pd.DataFrame(final_data, columns=['Ref', 'Design', 'Peca', 'Qtt', 'Unidade', 'preco', 'peso'])
            
            # Formatação do nome do ficheiro gerado
            nome_original = os.path.splitext(arquivo_laser.name)[0]
            nome_novo_ficheiro = f"{nome_original}_Importação_PHC.xlsx"
            
            buffer = io.BytesIO()
            df_final.to_excel(buffer, index=False, engine='xlsxwriter')
            
            # Guardamos o ficheiro processado na nossa lista
            arquivos_processados[nome_novo_ficheiro] = buffer.getvalue()
            
            # Mostramos no ecrã que este ficheiro específico já foi
            with st.expander(f"✅ Visualizar dados processados: {nome_novo_ficheiro}"):
                st.dataframe(df_final, use_container_width=True)
    
    # --- 3. LÓGICA DE DOWNLOAD (ÚNICO VS MÚLTIPLO) ---
    if len(arquivos_processados) == 1:
        # Se for só 1 ficheiro, fazemos o download direto do .xlsx
        nome_ficheiro, dados = list(arquivos_processados.items())[0]
        st.download_button("💾 Descarregar para PHC", dados, nome_ficheiro)
        
    elif len(arquivos_processados) > 1:
        # Se for mais de 1, criamos o pacote .zip
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for nome_ficheiro, dados in arquivos_processados.items():
                zip_file.writestr(nome_ficheiro, dados)
                
        st.download_button("🗂️ Descarregar Todos (Pacote ZIP)", zip_buffer.getvalue(), "Importacoes_PHC.zip", mime="application/zip")
