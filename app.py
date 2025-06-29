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
    df = df_input.copy()  # Trabalhar com uma cópia para não modificar o DataFrame original
    
    # Mapeamento de colunas esperadas (com nomes exatos confirmados)
    colunas_esperadas = [
        'Nome da reunião', 'Data de início da reunião', 'Data de término da reunião', 
        'Nome de exibição', 'Nome', 'Sobrenome', 'Função', 'E-mail do convidado', 
        'Hora da entrada', 'Hora da saída', 'Duração da presença', 
        'Tipo de conexão', 'Nome da sessão'
    ]

    # 1. Atribuir os nomes de colunas esperados (já foram lidos sem cabeçalho)
    # A verificação de colunas agora é feita pelo índice, pois a leitura é manual
    if len(df.columns) != len(colunas_esperadas):
        st.error(f"Erro: O número de colunas encontradas ({len(df.columns)}) não corresponde ao número de colunas esperadas ({len(colunas_esperadas)}).")
        st.info("Isso pode indicar que o arquivo tem um formato diferente ou dados corrompidos.")
        return None, None
        
    df.columns = colunas_esperadas # Atribui o cabeçalho correto
    
    # ===============================================================
    # --- PASSO DE DEPURAÇÃO: EXIBIR AS COLUNAS Lidas ---
    # ===============================================================
    st.info(f"Colunas do DataFrame (após atribuição manual): {list(df.columns)}")
    # ===============================================================

    # Contar registros totais e válidos antes da limpeza
    total_registros_processados = len(df)
    registros_validos_antes = len(df)
    
    # 2. Remover registros com dados faltantes essenciais
    df.dropna(subset=['E-mail do convidado', 'Hora da entrada', 'Hora da saída'], inplace=True)
    registros_ignorados = registros_validos_antes - len(df)
    
    # --- CORREÇÃO: CONVERTER COLUNA 'DURAÇÃO' PARA NÚMERO ---
    try:
        # Substituir vírgulas por pontos para o pandas entender como decimal
        df['Duração da presença'] = df['Duração da presença'].astype(str).str.replace(',', '.', regex=False)
        # Converter para numérico, forçando valores inválidos para NaN
        df['Duração da presença'] = pd.to_numeric(df['Duração da presença'], errors='coerce')
        # Remover linhas onde a duração não é um número
        registros_validos_antes_duracao = len(df)
        df.dropna(subset=['Duração da presença'], inplace=True)
        registros_ignorados += registros_validos_antes_duracao - len(df)
    except Exception as e:
        st.error(f"Erro ao limpar a coluna 'Duração da presença'. Verifique se ela contém apenas valores numéricos. Erro: {e}")
        return None, None
    # --------------------------------------------------------

    if df.empty:
        return None, {"total_registros_processados": total_registros_processados, 
                      "registros_ignorados": registros_ignorados, 
                      "presentes": 0, "ausentes": 0}

    try:
        # 3. Converter colunas de tempo para o formato datetime
        df['Hora da entrada'] = pd.to_datetime(df['Hora da entrada'])
        df['Hora da saída'] = pd.to_datetime(df['Hora da saída'])
        df['Data de início da reunião'] = pd.to_datetime(df['Data de início da reunião'])
        df['Data de término da reunião'] = pd.to_datetime(df['Data de término da reunião'])
    except Exception as e:
        st.error(f"Erro ao converter colunas de data/hora. Verifique o formato dos dados. Erro: {e}")
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

    # 5. Agrupar por e-mail para consolidar os registros de cada aluno
    grupos_por_email = df.groupby('E-mail do convidado')
    
    resultados = []

    # 6. Iterar sobre cada grupo (aluno) para consolidar os dados
    for email, grupo in grupos_por_email:
        entrada_consolidada = grupo['Hora da entrada'].min()
        saida_consolidada = grupo['Hora da saída'].max()
        tempo_total_min = grupo['Duração da presença'].sum()
        porcentagem_tempo = (tempo_total_min / duracao_total_aula_min) * 100
        
        # 7. Análise por tramos (slots) de 60 minutos
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
        status = 'Presente' if porcentagem_tempo >= 80 and porcentagem_tramos >= 80 else 'Ausente'
        
        try:
            nome_aluno = grupo.iloc[0]['Nome'] + ' ' + grupo.iloc[0]['Sobrenome']
        except KeyError:
            nome_aluno = grupo.iloc[0]['Nome de exibição']
            
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
    presentes = len(df_final[df_final['Status'] == 'Presente'])
    ausentes = len(df_final[df_final['Status'] == 'Ausente'])
    
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
        
        # Leemos el contenido del archivo como texto para el análisis manual
        file_content_bytes = uploaded_file.getvalue()
        decoded_content = None
        for encoding in ['utf-16', 'utf-8-sig', 'utf-8', 'latin1', 'cp1252']:
            try:
                decoded_content = file_content_bytes.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        
        if decoded_content is None:
            st.error("Não foi possível decodificar o arquivo. Tente salvá-lo em UTF-8.")
        else:
            # Separamos el contenido en líneas para procesar el encabezado
            lines = decoded_content.splitlines()
            if not lines:
                st.error("O arquivo está vazio.")
            else:
                # La primera línea contiene el encabezado. La segunda contiene los datos.
                header_line = lines[0]
                data_lines = lines[1:]
                
                # Leemos los datos a partir de la segunda línea, sin encabezado
                df_input = pd.read_csv(io.StringIO("\n".join(data_lines)), header=None, encoding=None, sep='\t')
                
                # Limpiamos el encabezado de BOMs y espacios y lo usamos para nombrar las columnas
                cleaned_header = [col.replace('\ufeff', '').strip() for col in header_line.split('\t')]
                df_input.columns = cleaned_header
                
                # Removemos la fila vacía extra que a veces viene después del encabezado
                df_input.dropna(how='all', inplace=True)
                
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
