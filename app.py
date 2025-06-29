import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import re

# ====================================================================
# Función de procesamiento del CSV de Webex
# ====================================================================
def processar_assistencia(df_input):
    """
    Processa um DataFrame de lista de presença do Webex e gera um relatório.
    """
    df = df_input.copy()
    
    # Mapeamento de colunas esperadas
    colunas_esperadas = [
        'Nome da reunião', 'Data de início da reunião', 'Data de término da reunião', 
        'Nome de exibição', 'Nome', 'Sobrenome', 'Função', 'E-mail do convidado', 
        'Hora da entrada', 'Hora da saída', 'Duração da presença', 
        'Tipo de conexão', 'Nome da sessão'
    ]

    # Atribuição e validação de colunas
    if len(df.columns) != len(colunas_esperadas):
        st.error(f"Erro: O número de colunas lidas ({len(df.columns)}) não corresponde ao esperado ({len(colunas_esperadas)}).")
        st.info("Isso pode indicar um formato de arquivo ou cabeçalho incorreto.")
        return None, None
        
    df.columns = colunas_esperadas
    
    st.info(f"Colunas do DataFrame (após atribuição manual): {list(df.columns)}")

    total_registros_processados = len(df)
    
    st.info(f"Total de registros lidos: {total_registros_processados}")
    
    registros_validos_antes = len(df)
    
    # 1. Remover registros com dados faltantes essenciais
    df.dropna(subset=['E-mail do convidado', 'Hora da entrada', 'Hora da saída', 'Duração da presença'], inplace=True)
    registros_ignorados = registros_validos_antes - len(df)

    st.info(f"Registros restantes após remover linhas com dados cruciais faltantes: {len(df)}")

    # 2. Converter a coluna de duração para numérico
    try:
        df['Duração da presença'] = df['Duração da presença'].astype(str)
        df['Duração da presença'] = df['Duração da presença'].str.replace(r'[^0-9,.]', '', regex=True)
        df['Duração da presença'] = df['Duração da presença'].str.replace(',', '.', regex=False)
        df['Duração da presença'] = pd.to_numeric(df['Duração da presença'], errors='coerce')
        
        st.info(f"Amostra da coluna 'Duração da presença' após a conversão: {list(df['Duração da presença'].head(5))}")
        
        registros_validos_antes_duracao = len(df)
        df.dropna(subset=['Duração da presença'], inplace=True)
        registros_ignorados += registros_validos_antes_duracao - len(df)
        
        st.info(f"Registros restantes após limpar a duração: {len(df)}")
        
    except Exception as e:
        st.error(f"Erro ao limpar a coluna 'Duração da presença'. Erro: {e}")
        return None, None

    if df.empty:
        st.warning("O DataFrame está vazio após a limpeza dos dados.")
        return None, {"total_registros_processados": total_registros_processados, 
                      "registros_ignorados": registros_ignorados, 
                      "presentes": 0, "ausentes": 0}

    # 3. Converter colunas de tempo para o formato datetime
    try:
        st.info("Tentando converter colunas de data/hora...")
        # REMOVER FORMATO DE FÓRMULA DO EXCEL
        for col in ['Hora da entrada', 'Hora da saída', 'Data de início da reunião', 'Data de término da reunião']:
            df[col] = df[col].astype(str).str.replace('="', '', regex=False).str.replace('"', '', regex=False)
        
        df['Hora da entrada'] = pd.to_datetime(df['Hora da entrada'])
        df['Hora da saída'] = pd.to_datetime(df['Hora da saída'])
        df['Data de início da reunião'] = pd.to_datetime(df['Data de início da reunião'])
        df['Data de término da reunião'] = pd.to_datetime(df['Data de término da reunião'])
    except Exception as e:
        st.error(f"Erro ao converter colunas de data/hora. Verifique o formato dos dados. Erro: {e}")
        st.warning(f"Amostra de dados que causou o erro: {list(df['Hora da entrada'].head())}")
        return None, None

    # 4. Calcular a duração total da aula
    try:
        duracao_total_aula_min = (df['Data de término da reunião'].iloc[0] - df['Data de início da reunião'].iloc[0]).total_seconds() / 60
    except IndexError:
        st.error("Erro: O arquivo está vazio ou não contém informações de duração da reunião.")
        return None, None

    if duracao_total_aula_min <= 0:
        st.error("Erro: Não foi possível calcular a duração da aula. Verifique as datas de início e término da reunião.")
        return None, None

    # GARANTIR QUE NOME/SOBRENOME SEJAM STRINGS
    df['Nome'] = df['Nome'].fillna('').astype(str)
    df['Sobrenome'] = df['Sobrenome'].fillna('').astype(str)

    # 5. Agrupar por e-mail para consolidar os registros
    grupos_por_email = df.groupby('E-mail do convidado')
    
    resultados = []
    
    # 6. Iterar sobre cada grupo (aluno)
    for email, grupo in grupos_por_email:
        entrada_consolidada = grupo['Hora da entrada'].min()
        saida_consolidada = grupo['Hora da saída'].max()
        tempo_total_min = grupo['Duração da presença'].sum()
        porcentagem_tempo = (tempo_total_min / duracao_total_aula_min) * 100
        
        # 7. Análise por tramos
        tramos_participados = 0
        total_tramos = int(duracao_total_aula_min / 60)
        if duracao_total_aula_min % 60 > 0:
            total_tramos += 1
        
        hora_inicio_aula = df['Data de início da reunião'].iloc[0]
        
        for i in range(total_tramos):
            inicio_tramo = hora_inicio_aula + timedelta(minutes=i*60)
            fim_tramo = inicio_tramo + timedelta(minutes=60)
            participou_do_tramo = False
            for _, registro in grupo.iterrows():
                if (registro['Hora da entrada'] < fim_tramo) and (registro['Hora da saída'] > inicio_tramo):
                    participou_do_tramo = True
                    break
            if participou_do_tramo:
                tramos_participados += 1

        porcentagem_tramos = (tramos_participados / total_tramos) * 100 if total_tramos > 0 else 0
        
        # --- CORREÇÃO FINAL: Substituir status por P e FI ---
        status = 'P' if porcentagem_tempo >= 80 and porcentagem_tramos >= 80 else 'FI'
        # -----------------------------------------------------
        
        nome_aluno = str(grupo.iloc[0]['Nome']) + ' ' + str(grupo.iloc[0]['Sobrenome'])
            
        resultados.append({
            'Nome': nome_aluno,
            'E-mail do convidado': email,
            'Entrada Consolidada': entrada_consolidada.strftime('%Y-%m-%d %H:%M:%S'),
            'Saída Consolidada': saida_consolidada.strftime('%Y-%m-%d %H:%M:%S'),
            'Tempo Total (min)': round(tempo_total_min, 2),
            'Porcentagem de Tempo (%)': round(porcentagem_tempo, 2),
            'Tramos Participados': tramos_participados,
            'Porcentagem de Tramos (%)': round(porcentagem_tramos, 2),
            'Status': status
        })

    # 8. Gerar o novo DataFrame
    df_final = pd.DataFrame(resultados)
    
    # 9. Gerar o resumo
    presentes = len(df_final[df_final['Status'] == 'P'])
    ausentes = len(df_final[df_final['Status'] == 'FI'])
    
    resumo = {
        "total_registros_processados": total_registros_processados,
        "registros_ignorados": registros_ignorados,
        "presentes": presentes,
        "ausentes": ausentes
    }
    
    return df_final, resumo

# ====================================================================
# Interfaz de usuario con Streamlit
# ====================================================================

st.set_page_config(page_title="Procesador de Asistencia Webex", layout="wide", page_icon="👨‍🏫")
st.title("👨‍🏫 Procesador de Asistencia para Moodle")
st.markdown("Suba su archivo CSV de la lista de presencia de Webex para generar un informe listo para importar en Moodle.")
st.divider()

uploaded_file = st.file_uploader("📥 Cargue el archivo CSV aquí", type=["csv"])

if uploaded_file is not None:
    try:
        df_input = None
        read_configs = [
            {'encoding': 'utf-16', 'sep': '\t', 'header': 0},
            {'encoding': 'utf-8-sig', 'sep': '\t', 'header': 0}, 
            {'encoding': 'latin1', 'sep': '\t', 'header': 0},
            {'encoding': 'utf-8', 'sep': '\t', 'header': 0},
            {'encoding': 'utf-8', 'delim_whitespace': True, 'header': 0},
            {'encoding': 'latin1', 'delim_whitespace': True, 'header': 0},
            {'encoding': 'utf-8', 'sep': ',', 'header': 0},
            {'encoding': 'latin1', 'sep': ',', 'header': 0},
            {'encoding': 'utf-8', 'sep': ';', 'header': 0},
            {'encoding': 'latin1', 'sep': ';', 'header': 0},
        ]
        
        for config in read_configs:
            try:
                uploaded_file.seek(0)
                df_input = pd.read_csv(uploaded_file, **config)
                if not df_input.empty and len(df_input.columns) > 1:
                    st.info(f"Arquivo lido com sucesso! Delimitador: '{config.get('sep', 'whitespace')}', Codificação: '{config['encoding']}'.")
                    break  
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
            except Exception:
                continue

        if df_input is None or df_input.empty or len(df_input.columns) <= 1:
            st.error("Erro ao ler o arquivo. Não foi possível determinar a codificação ou o delimitador correto. Tente salvar o CSV como UTF-8 com vírgulas ou tabulações como delimitador.")
        
        else:
            st.success("¡Archivo cargado con éxito!")
            st.info("Procesando los datos... por favor, espere.")
    
            df_reporte, resumen_final = processar_assistencia(df_input)
    
            if df_reporte is not None:
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("👥 Registros Totales", resumen_final['total_registros_processados'])
                col2.metric("❌ Registros Ignorados", resumen_final['registros_ignorados'])
                col3.metric("✅ Estudiantes Presentes", resumen_final['presentes'])
                col4.metric("🚫 Estudiantes Ausentes", resumen_final['ausentes'])
                
                st.divider()
                st.header("📊 Reporte Final de Asistencia")
                st.dataframe(df_reporte, use_container_width=True)
    
                csv_buffer = io.StringIO()
                df_reporte.to_csv(csv_buffer, index=False, encoding='utf-8')
                csv_bytes = csv_buffer.getvalue().encode('utf-8')
    
                st.download_button(
                    label="📤 Descargar Reporte CSV",
                    data=csv_bytes,
                    file_name="reporte_asistencia_moodle.csv",
                    mime="text/csv",
                    help="Clique para baixar o arquivo CSV final."
                )
            else:
                st.warning("No se pudo generar el reporte. Verifique el formato de su archivo CSV.")

    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")
        st.info("Asegúrese de que el archivo es un CSV válido de Webex y de que tiene todas las columnas requeridas.")

st.divider()
st.markdown("Creado con ❤️ por el Agente Procesador de Asistencia.")
