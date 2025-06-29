import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# ====================================================================
# Función de procesamiento del CSV de Webex
# ====================================================================
def processar_assistencia(df):
    """
    Processa um DataFrame de lista de presença do Webex e gera um relatório.
    """
    # Contar registros totais y válidos antes de la limpieza
    total_registros_processados = len(df)
    registros_validos_antes = len(df)
    
    # Remover registros con datos faltantes esenciales
    df.dropna(subset=['E-mail do convidado', 'Hora da entrada', 'Hora da saída'], inplace=True)
    registros_ignorados = registros_validos_antes - len(df)
    
    if df.empty:
        return None, {"total_registros_processados": total_registros_processados, 
                      "registros_ignorados": registros_ignorados, 
                      "presentes": 0, "ausentes": 0}

    try:
        # Convertir colunas de tiempo para el formato datetime
        df['Hora da entrada'] = pd.to_datetime(df['Hora da entrada'])
        df['Hora da saída'] = pd.to_datetime(df['Hora da saída'])
        df['Data de início da reunião'] = pd.to_datetime(df['Data de início da reunião'])
        df['Data de término da reunião'] = pd.to_datetime(df['Data de término da reunião'])
    except KeyError as e:
        st.error(f"Erro: Coluna '{e.args[0]}' não encontrada no arquivo CSV. Certifique-se de que o arquivo é um relatório Webex válido com todas as colunas esperadas.")
        return None, None
    except Exception as e:
        st.error(f"Erro ao converter colunas de data/hora: {e}. Verifique o formato dos dados.")
        return None, None

    # Calcular a duração total da aula
    try:
        # Usamos .iloc[0] para pegar o primeiro registro, assumindo que a duração é a mesma para todos
        duracao_total_aula_min = (df['Data de término da reunião'].iloc[0] - df['Data de início da reunião'].iloc[0]).total_seconds() / 60
    except IndexError:
        st.error("Erro: O arquivo está vazio ou não contém informações de duração da reunião.")
        return None, None

    if duracao_total_aula_min <= 0:
        st.error("Erro: Não foi possível calcular a duração da aula. Verifique as datas de início e término da reunião.")
        return None, None

    # Agrupar por e-mail para consolidar os registros de cada aluno
    grupos_por_email = df.groupby('E-mail do convidado')
    
    # Dicionário para armazenar os resultados consolidados
    resultados = []

    # Iterar sobre cada grupo (aluno) para consolidar os dados
    for email, grupo in grupos_por_email:
        # Consolidar entradas e saídas
        entrada_consolidada = grupo['Hora da entrada'].min()
        saida_consolidada = grupo['Hora da saída'].max()
        
        # Somar o tempo total de presença em minutos
        tempo_total_min = grupo['Duração da presença'].sum()

        # Calcular a porcentagem de tempo
        porcentagem_tempo = (tempo_total_min / duracao_total_aula_min) * 100
        
        # Análise por tramos (slots) de 60 minutos
        tramos_participados = 0
        total_tramos = int(duracao_total_aula_min / 60)
        if duracao_total_aula_min % 60 > 0:
            total_tramos += 1
        
        hora_inicio_aula = df['Data de início da reunião'].iloc[0]
        
        for i in range(total_tramos):
            inicio_tramo = hora_inicio_aula + timedelta(minutes=i*60)
            fim_tramo = inicio_tramo + timedelta(minutes=60)
            
            # Verificar se o aluno participou do tramo
            participou_do_tramo = False
            for _, registro in grupo.iterrows():
                # Verifica se o intervalo de tempo do tramo se sobrepõe ao tempo de presença do aluno
                if (registro['Hora da entrada'] < fim_tramo) and (registro['Hora da saída'] > inicio_tramo):
                    participou_do_tramo = True
                    break
            
            if participou_do_tramo:
                tramos_participados += 1

        porcentagem_tramos = (tramos_participados / total_tramos) * 100 if total_tramos > 0 else 0

        # Determinar o status
        status = 'Presente' if porcentagem_tempo >= 80 and porcentagem_tramos >= 80 else 'Ausente'
        
        # Obter nome do aluno (pode pegar o primeiro registro)
        try:
            nome_aluno = grupo.iloc[0]['Nome'] + ' ' + grupo.iloc[0]['Sobrenome']
        except KeyError:
            nome_aluno = grupo.iloc[0]['Nome de exibição'] # Usa o nome de exibição se Nome/Sobrenome faltarem
            
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

    # Gerar o novo DataFrame
    df_final = pd.DataFrame(resultados)
    
    # Gerar o resumo
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
# Tenta ler com a codificação padrão 'utf-8' primeiro
try:
    df_input = pd.read_csv(io.BytesIO(uploaded_file.getvalue()), encoding='utf-8')
except UnicodeDecodeError:
    # Se falhar, tenta com a codificação 'latin1'
    try:
        df_input = pd.read_csv(io.BytesIO(uploaded_file.getvalue()), encoding='latin1')
    except Exception as e:
        # Se falhar novamente, exibe um erro mais genérico
        st.error(f"Erro de codificação. Tente salvar o arquivo CSV com a codificação UTF-8. Erro detalhado: {e}")
        return
        
        st.success("¡Archivo cargado con éxito!")
        st.info("Procesando los datos... por favor, espere.")

        # Llamar a la función de procesamiento
        df_reporte, resumen_final = processar_assistencia(df_input.copy())

        if df_reporte is not None:
            # Mostrar el resumen en columnas
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("👥 Registros Totales", resumen_final['total_registros_processados'])
            col2.metric("❌ Registros Ignorados", resumen_final['registros_ignorados'])
            col3.metric("✅ Estudiantes Presentes", resumen_final['presentes'])
            col4.metric("🚫 Estudiantes Ausentes", resumen_final['ausentes'])
            
            st.divider()
            st.header("📊 Reporte Final de Asistencia")
            st.dataframe(df_reporte, use_container_width=True)

            # Crear el link de descarga
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
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
        st.info("Asegúrese de que el archivo es un CSV válido de Webex y de que tiene todas las columnas requeridas (por ejemplo, 'Data de início da reunião', 'E-mail do convidado', etc.).")

st.divider()
st.markdown("Creado con ❤️ por el Agente Procesador de Asistencia.")
