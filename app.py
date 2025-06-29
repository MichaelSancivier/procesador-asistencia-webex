Ah, entendi perfeitamente\! Receber duas mensagens de erro diferentes indica que seu arquivo CSV tem dois problemas: um de **codificação** e outro de **formato/nome das colunas**.

Vamos analisar cada um e ver a solução.

-----

### **Análise dos Erros e Soluções**

#### **Erro 1: `Erro: A coluna 'Nome da reunião' não foi encontrada no arquivo.`**

Este erro é muito claro. O agente está procurando uma coluna com o nome `Nome da reunião`, mas não a encontrou no seu arquivo. Isso pode acontecer por vários motivos:

  * **Diferença de idioma:** O relatório Webex foi gerado em inglês ou em outro idioma, e o nome da coluna é `Meeting Name` em vez de `Nome da reunião`.
  * **Nome da coluna ligeiramente diferente:** Pode ser que o nome seja `Nome da Reunião` (com "R" maiúsculo) ou `NomeReunião` sem espaços.
  * **Espaços extras:** O nome pode ter um espaço extra no início ou no fim (`" Nome da reunião "` ou `"Nome da reunião "`).
  * **Relatório diferente:** Você pode ter baixado um tipo de relatório diferente (por exemplo, um relatório de uso em vez de um relatório de presença detalhado), que não contém todas as colunas que o agente espera.

**Solução:**

Você precisa garantir que o nome da coluna no seu arquivo CSV seja **exatamente** o que o código espera. O código que te enviei agora faz uma verificação inicial e vai te dizer qual coluna está faltando, o que é ótimo.

Para resolver, você tem duas opções:

1.  **Ajuste o nome da coluna no seu CSV:** A maneira mais simples. Abra o arquivo CSV em um editor de texto (como Notepad++), Google Sheets ou Excel e mude o nome do cabeçalho da coluna para `Nome da reunião` (e os outros nomes que ele possa reclamar depois). Salve o arquivo.
2.  **Ajuste o código:** Se você sempre recebe relatórios com nomes de coluna diferentes, podemos ajustar o código para ser mais flexível, mapeando os nomes. Por exemplo:
    ```python
    df.rename(columns={'Meeting Name': 'Nome da reunião', 'Duration of presence': 'Duração da presença'}, inplace=True)
    ```

-----

#### **Erro 2: `No se pudo generar el reporte. Verifique el formato de su archivo CSV.`**

Este erro é o resultado do primeiro. Como a função `processar_assistencia` não consegue encontrar as colunas essenciais, ela retorna `None`, e a interface do Streamlit exibe essa mensagem genérica de que o relatório não pôde ser gerado.

**Solução:**

A solução para este erro depende de você resolver o **Erro 1**. Assim que o agente conseguir encontrar todas as colunas que precisa, o relatório será gerado normalmente.

-----

### **Recomendação e Próximos Passos**

1.  **Abra o arquivo CSV original:** Use um editor de texto simples para abrir o arquivo CSV exportado do Webex.
2.  **Verifique os nomes das colunas:** Olhe atentamente a primeira linha do arquivo (o cabeçalho) e veja os nomes exatos das colunas.
3.  **Ajuste seu código `app.py`:** Use a lista de colunas que você encontrou para ajustar a verificação no código. Eu vou te dar uma versão do código que já faz isso de forma mais robusta, mas você precisará verificar os nomes.

Aqui está o código `app.py` com uma verificação de colunas mais robusta e a correção de codificação:

```python
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
    # Contar registros totais e válidos antes da limpeza
    total_registros_processados = len(df)
    
    # Mapeamento de colunas esperadas
    colunas_esperadas = {
        'Nome da reunião': 'Nome da reunião',
        'Data de início da reunião': 'Data de início da reunião',
        'Data de término da reunião': 'Data de término da reunião',
        'Nome de exibição': 'Nome de exibição',
        'Nome': 'Nome',
        'Sobrenome': 'Sobrenome',
        'Função': 'Função',
        'E-mail do convidado': 'E-mail do convidado',
        'Hora da entrada': 'Hora da entrada',
        'Hora da saída': 'Hora da saída',
        'Duração da presença': 'Duração da presença',
        'Tipo de conexão': 'Tipo de conexão',
        'Nome da sessão': 'Nome da sessão'
    }

    # Verificar se todas as colunas esperadas estão presentes, ignorando a caixa (maiúsculas/minúsculas)
    colunas_df = {col.strip().lower(): col for col in df.columns}
    colunas_faltantes = []
    
    for coluna_esperada in colunas_esperadas.keys():
        if coluna_esperada.lower() not in colunas_df:
            colunas_faltantes.append(coluna_esperada)
    
    if colunas_faltantes:
        st.error(f"Erro: As seguintes colunas não foram encontradas no arquivo: {', '.join(colunas_faltantes)}.")
        st.info("Verifique se o arquivo CSV é um relatório de presença Webex válido e se as colunas estão nomeadas corretamente.")
        return None, None

    # Normalizar os nomes das colunas para os nomes esperados
    df.columns = [colunas_esperadas.get(c.lower(), c) for c in df.columns]

    registros_validos_antes = len(df)
    
    # Remover registros com dados faltantes essenciais
    df.dropna(subset=['E-mail do convidado', 'Hora da entrada', 'Hora da saída'], inplace=True)
    registros_ignorados = registros_validos_antes - len(df)
    
    if df.empty:
        return None, {"total_registros_processados": total_registros_processados, 
                      "registros_ignorados": registros_ignorados, 
                      "presentes": 0, "ausentes": 0}

    try:
        # Converter colunas de tempo para o formato datetime
        df['Hora da entrada'] = pd.to_datetime(df['Hora da entrada'])
        df['Hora da saída'] = pd.to_datetime(df['Hora da saída'])
        df['Data de início da reunião'] = pd.to_datetime(df['Data de início da reunião'])
        df['Data de término da reunião'] = pd.to_datetime(df['Data de término da reunião'])
    except Exception as e:
        st.error(f"Erro ao converter colunas de data/hora. Verifique o formato dos dados. Erro: {e}")
        return None, None

    # Calcular a duração total da aula
    try:
        duracao_total_aula_min = (df['Data de término da reunião'].iloc[0] - df['Data de início da reunião'].iloc[0]).total_seconds() / 60
    except IndexError:
        st.error("Erro: O arquivo está vazio ou não contém informações de duração da reunião.")
        return None, None

    if duracao_total_aula_min <= 0:
        st.error("Erro: Não foi possível calcular a duração da aula. Verifique as datas de início e término da reunião.")
        return None, None

    # Agrupar por e-mail para consolidar os registros de cada aluno
    grupos_por_email = df.groupby('E-mail do convidado')
    
    resultados = []

    # Iterar sobre cada grupo (aluno) para consolidar os dados
    for email, grupo in grupos_por_email:
        entrada_consolidada = grupo['Hora da entrada'].min()
        saida_consolidada = grupo['Hora da saída'].max()
        tempo_total_min = grupo['Duração da presença'].sum()
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

    df_final = pd.DataFrame(resultados)
    
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
        # AQUI ESTÁ A CORREÇÃO DE CODIFICAÇÃO E A LEITURA DO CSV
        try:
            df_input = pd.read_csv(uploaded_file, encoding='utf-8')
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df_input = pd.read_csv(uploaded_file, encoding='latin1')
        except pd.errors.ParserError as e:
            st.error(f"Erro ao analisar o arquivo CSV. Verifique se ele está bem formatado (ex: vírgulas separando os dados). Erro: {e}")
            df_input = None
            
        if df_input is not None:
            st.success("¡Archivo cargado con éxito!")
            st.info("Procesando los datos... por favor, espere.")
    
            # Chamar a função de processamento
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
        st.error(f"Ocurrió un error inesperado: {e}")
        st.info("Asegúrese de que el archivo es un CSV válido de Webex y de que tiene todas las columnas requeridas (por ejemplo, 'Data de início da reunião', 'E-mail do convidado', etc.).")

st.divider()
st.markdown("Creado con ❤️ por el Agente Procesador de Asistencia.")
```
