import pandas as pd
import streamlit as st

# Carregar os dados
df = pd.read_excel("sat_rawdataweek.xlsx")

# Garantir que as colunas relevantes sejam do tipo numérico
df['AHT'] = pd.to_numeric(df['AHT'], errors='coerce').astype(float)
df['Target_Effec Work Hours'] = pd.to_numeric(df['Target_Effec Work Hours'], errors='coerce').astype(float)

# LINHAS USADAS COM OS DADOS NOVOS RAWDATA2
#df['WFM_Saturation %'] = df['WFM_Saturation %'].str.replace('%', '').astype(float) / 100
#df['WFM_Saturation %'] = df['WFM_Saturation %'].fillna(0)
#print(df['WFM_Saturation %'])

df['WFM_Saturation %'] = pd.to_numeric(df['WFM_Saturation %'], errors='coerce')
df['New Saturation'] = (df['Total Volume'] * df['AHT'] / (3600 * df['Target_Effec Work Hours']) / df['HC']) * 100

# Carregar os dados na sessão
if 'df' not in st.session_state:
    st.session_state.df = df

# Inicializar as variáveis de controle
if 'show_create_form' not in st.session_state:
    st.session_state.show_create_form = False
if 'show_delete_form' not in st.session_state:
    st.session_state.show_delete_form = False

# Função para calcular a nova saturação
def calculate_new_saturation(aht, target_ewh, actual_fte, selected_volume):
    if actual_fte == 0 or target_ewh == 0:
        return None  # Evitar divisão por zero
    
    new_saturation = ((selected_volume * (aht / 3600) / target_ewh) / actual_fte) * 100
    return new_saturation

# Função para atualizar o DataFrame com a nova saturação
def update_dataframe_with_new_saturation(df, volumes_dict, selected_market, selected_week):
    df_copy = df[(df['MARKET_LANGUAGE_L5'] == selected_market) & (df['WEEK_Monday_name'] == selected_week) & (df['LOB_L1'] == selected_lob) & (df['LOB_L2'] == selected_sublob)].copy()

    for selected_queue, selected_volume in volumes_dict.items():
        if selected_queue in df_copy['QUEUE_NAME_L8'].values:
            filtered_df = df_copy[df_copy['QUEUE_NAME_L8'] == selected_queue]
            aht = filtered_df['AHT'].sum()
            target_ewh = filtered_df['Target_Effec Work Hours'].sum()
            actual_fte = filtered_df['HC'].sum()
            new_saturation = calculate_new_saturation(aht, target_ewh, actual_fte, selected_volume)

            # Aplicar o cálculo apenas à fila selecionada
            df_copy.loc[df_copy['QUEUE_NAME_L8'] == selected_queue, 'New Saturation'] = new_saturation

    # Garantir que a coluna 'New Saturation' contenha valores numéricos
    df_copy['New Saturation'] = df_copy['New Saturation'].fillna(df_copy['WFM_Saturation %'] * 100)

    return df_copy

# Função para adicionar subtotais
def add_subtotals(df):
    subtotal_aht = df['AHT'].sum()
    subtotal_target = df['Target_Effec Work Hours'].sum()
    subtotal_wfm_saturation = df['WFM_Saturation %'].sum()
    subtotal_volume = df["Total Volume"].sum()
    subtotal_new_saturation = df['New Saturation'].sum()

    subtotal_row = pd.DataFrame([{
        'QUEUE_NAME_L8': 'Subtotal',
        'AHT': subtotal_aht,
        'Target_Effec Work Hours': subtotal_target,
        'WFM_Saturation %': subtotal_wfm_saturation,
        'New Saturation': subtotal_new_saturation,
        'Total Volume': subtotal_volume,
        'WEEK_Monday_name': selected_week
    }])
    
    df_with_subtotal = pd.concat([df, subtotal_row], ignore_index=True)
    return df_with_subtotal

# Função para mudar a cor do background
def highlight_selected_queues(row, selected_queues):
    styles = [''] * len(row)
    if row['QUEUE_NAME_L8'] in selected_queues:
        index = row.index.get_loc('New Saturation')
        index_queue_name = row.index.get_loc('QUEUE_NAME_L8')
        styles[index] = 'background-color: darkred'
        styles[index_queue_name] = 'background-color: darkred'
    return styles

# Função para adicionar uma nova fila ao DataFrame
def add_new_queue(df, queue_name, aht, target_ewh, total_volume, hc, market, week):
    new_saturation = calculate_new_saturation(aht, target_ewh, hc, total_volume)
    new_row = pd.DataFrame([{
        'QUEUE_NAME_L8': queue_name,
        'AHT': aht,
        'Target_Effec Work Hours': target_ewh,
        'WFM_Saturation %': 0,
        'New Saturation': new_saturation,
        'Total Volume': total_volume,
        'HC': hc,
        "WEEK_Monday_name": week,
        'MARKET_LANGUAGE_L5': market,
        'LOB_L1' : lob,
        'LOB_L2' : sublob
    }])
    st.session_state.df = pd.concat([df, new_row], ignore_index=True)
   
# Função para deletar uma fila do DataFrame
def delete_queue(df, queue_name):
    df = df[df['QUEUE_NAME_L8'] != queue_name].reset_index(drop=True)  # Resetar o índice após a exclusão
    st.session_state.df = df

# Página
st.set_page_config(layout="wide")
st.title("Saturation Rate Prediction")
st.subheader("Select LOB and Market to review")

# Select LOBs
with st.container():
    col1, col2, col3 = st.columns(3)

    with col1:
        selected_lob = st.selectbox("LOB", st.session_state.df["LOB_L1"].unique())

    with col2:
        selected_sublob = st.selectbox("sub-LOB", st.session_state.df[st.session_state.df["LOB_L1"] == selected_lob]["LOB_L2"].unique())
        # Filtrar o DataFrame com base nos LOBs selecionados
        df_filtered_lob = st.session_state.df[(st.session_state.df["LOB_L1"] == selected_lob) & (st.session_state.df["LOB_L2"] == selected_sublob)]
    
    # Selectbox Mercados
    with col3:
        selected_market = st.selectbox("Market", df_filtered_lob["MARKET_LANGUAGE_L5"].unique())

#st.divider()

col1, col2, col3 = st.columns(3)

# Selectbox Semanas
# Obtendo todas as semanas únicas
all_weeks = df["WEEK_Monday_name"].unique().astype(str)
all_weeks.sort()
with col2:
    selected_week = st.selectbox("Week", all_weeks)

# Filtrando o DataFrame com a Week e LOB
df_filtered_week = df_filtered_lob[df_filtered_lob["WEEK_Monday_name"] == selected_week]

# Criar um multiselect para selecionar as filas
with col1:
    filtered_df = df_filtered_week[(df_filtered_week["MARKET_LANGUAGE_L5"] == selected_market)]
    selected_queues = st.multiselect('Select Queues to edit volume', filtered_df['QUEUE_NAME_L8'].unique())

# Criar inputs numéricos para cada fila selecionada
volumes_dict = {}
for queue in selected_queues:
    volumes_dict[queue] = st.number_input(f'Set a Volume for {queue}', min_value=0, max_value=500000, value=10000)

# Botão para mostrar/ocultar o formulário de criação de nova fila
col1, col2, col3 = st.columns((1, 1, 4))

with col1:
    if st.button('Create New Queue'):
        st.session_state.show_create_form = not st.session_state.show_create_form

# Formulário para criar uma nova fila
if st.session_state.show_create_form:
    with st.form(key='new_queue_form'):
        queue_name = st.text_input('Queue Name')
        aht = st.number_input('AHT', min_value=0.0, step=1.0)
        target_ewh = filtered_df['Target_Effec Work Hours'].iloc[0]
        total_volume = st.number_input('Total Volume', min_value=0)
        hc = filtered_df['HC'].iloc[0]
        market = selected_market
        week = selected_week
        submitted = st.form_submit_button('Add Queue')
        lob = selected_lob
        sublob = selected_sublob
        
        if submitted:
            add_new_queue(st.session_state.df, queue_name, aht, target_ewh, total_volume, hc, market, week)
            st.success('New queue added successfully!')
            st.session_state.show_create_form = False  # Fechar o formulário
            st.experimental_rerun()  # Recarregar a página para fechar o formulário

# Botão para mostrar/ocultar o formulário de deletar fila
with col2:
    if st.button('Delete Queue'):
        st.session_state.show_delete_form = not st.session_state.show_delete_form

# Formulário para deletar uma fila
if st.session_state.show_delete_form:
    with st.form(key='delete_queue_form'):
        queue_to_delete = st.selectbox('Select Queue to Delete', filtered_df['QUEUE_NAME_L8'].unique())
        delete_submitted = st.form_submit_button('Delete Queue')
        
        if delete_submitted:
            delete_queue(st.session_state.df, queue_to_delete)
            st.success('Queue deleted successfully!')
            st.session_state.show_delete_form = False  # Fechar o formulário
            st.experimental_rerun()  # Recarregar a página para fechar o formulário

# Calcular a nova saturação e atualizar o DataFrame
df_updated = update_dataframe_with_new_saturation(st.session_state.df, volumes_dict, selected_market, selected_week)

# Adicionar os subtotais ao DataFrame
df_with_subtotal = add_subtotals(df_updated)

df_with_subtotal["WFM_Saturation %"] = df_with_subtotal["WFM_Saturation %"] * 100
df_with_subtotal = df_with_subtotal.dropna(subset=["Total Volume"])

# Exibir o DataFrame atualizado no Streamlit
col1, col2 = st.columns([6, 1])
with col1: 
    # Estilizar o DataFrame e aplicar formatação
    styled_df = (
        df_with_subtotal
        .style
        .format({
            'AHT': '{:.2f}',
            'Total Volume': '{:.0f}',
            'HC': '{:.2f}',
            'Target_Effec Work Hours': '{:.2f}',
            'New Saturation': '{:.2f}%',
            'WFM_Saturation %': '{:.2f}%'
        })
        .apply(lambda row: highlight_selected_queues(row, selected_queues), axis=1)
    )

    # Exibir o DataFrame estilizado
    st.dataframe(styled_df, height=500)

# Exibir card subtotal com New Saturation
with col2:
    subtotal_saturation = df_with_subtotal.iloc[-1]['WFM_Saturation %']
    st.metric(label="Actual Saturation", value=f"{subtotal_saturation:.2f}%")

    st.text("-------------------")

    subtotal_new_saturation = df_with_subtotal.iloc[-1]['New Saturation']
    st.metric(label="New Saturation", value=f"{subtotal_new_saturation:.2f}%", delta_color="inverse", delta=(subtotal_new_saturation - subtotal_saturation).round(2))

#st.divider()

st.title("Further Analysis")

# Adicionar gráfico de linhas
df_market = df[df["MARKET_LANGUAGE_L5"] == selected_market]

# Calcular os subtotais por semana
df_market_grouped = df_market.groupby("WEEK_Monday_name").agg({
    "WFM_Saturation %": "mean",
    "New Saturation": "mean"
}).reset_index()

# Ordenar as semanas corretamente e ajustar colunas
df_market_grouped = df_market_grouped.sort_values('WEEK_Monday_name')
df_market_grouped["WFM_Saturation %"] *=100
df_market_grouped["New Saturation"] = df_market_grouped["New Saturation"].round(2)
# Ajustando a saturação conforme modificações
for index, row in df_market_grouped.iterrows():
    if row["WEEK_Monday_name"] == selected_week:
        df_market_grouped.at[index, "New Saturation"] = subtotal_new_saturation.round(2)

# Verificar se os dados estão corretos
#st.write(df_market_grouped)

# Gráfico de linhas ajustado
#st.subheader("Adjusted Line Chart")
#st.line_chart(data=df_market_grouped, x= "WEEK_Monday_name", y=["New Saturation", "WFM_Saturation %"], use_container_width=True)
#st.write(df_market_grouped)