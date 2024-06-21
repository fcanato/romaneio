import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# Função para salvar DataFrame em um arquivo Excel em memória
def to_excel(df, cidade):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Adicionar título
        workbook  = writer.book
        worksheet = workbook.add_worksheet('Romaneio de Cargas')
        
        # Escrever o título na primeira linha, incluindo o nome da cidade
        titulo = f'Romaneio de Cargas Tel {cidade}'
        worksheet.merge_range('A1:I1', titulo, workbook.add_format({'align': 'center', 'bold': True, 'font_size': 14}))

        # Adicionar os dados do DataFrame começando da segunda linha
        df.to_excel(writer, index=False, sheet_name='Romaneio de Cargas', startrow=1)

        # Calcular a posição das assinaturas
        row_start = len(df) + 3  # Três linhas abaixo do último registro (incluindo o título e uma linha em branco)

        # Escrever campos de assinatura e data de saída
        worksheet.write(row_start, 0, "Conferente:")
        worksheet.write(row_start + 1, 0, "Motorista:")
        worksheet.write(row_start + 2, 0, "Data de Saída:")
        
    processed_data = output.getvalue()
    return processed_data

# Configuração da página
st.set_page_config(page_title="Romaneio de Cargas de Envio", page_icon="🚚", layout='wide', initial_sidebar_state="expanded")

def main():
    st.sidebar.header("Romaneio de Cargas Envio Tel")
    
    # Caminho da imagem na mesma pasta do arquivo Python
    image_path = os.path.join(os.path.dirname(__file__), "romaneio.jpg")
    st.sidebar.image(image_path, use_column_width=True)
    
    st.title("Gerenciamento de Romaneio de Cargas")

    # Carregar dados existentes
    def load_data():
        try:
            return pd.read_csv("romaneio_cargas.csv")
        except FileNotFoundError:
            return pd.DataFrame(columns=["Número Transferência", "Cidade Origem", "Cidade Destino", "Quantidade Volumes", "Conferente", "Motorista", "Data Saída", "Cidade Transbordo", "Destino Final"])

    # Função para salvar dados
    def save_data(df):
        df.to_csv("romaneio_cargas.csv", index=False)

    df = load_data()

    # Listas predefinidas de cidades
    cidades = ["Ribeirão Preto", "Araraquara", "Belo Horizonte", "São Paulo", "Bauru", "Presidente Prudente", "Araçatuba", "Jundiai", "São Jose do Rio Preto", "Marilia",
               "Piracicaba",'Sorocaba','Santa Barbara do Oeste','Cubatão','Rio de Janeiro']

    # Entradas do usuário
    st.header("Informações da Transferência")
    numero_transferencia = st.text_input("Número de Transferência SGM ou documento Vivo")
    cidade_origem = st.selectbox("Cidade de Origem", cidades)
    cidade_destino = st.selectbox("Cidade de Destino", cidades)
    transbordo = st.selectbox("Destino tem transbordo?", ["Não", "Sim"])
    
    cidade_transbordo = ""
    destino_final = ""

    if transbordo == "Sim":
        cidade_transbordo = st.selectbox("Cidade de Transbordo", cidades)
        destino_final = st.selectbox("Destino Final", cidades)
    else:
        destino_final = cidade_destino

    quantidade_volumes = st.number_input("Quantidade de Volumes", min_value=1, step=1)
    conferente = st.text_input("Conferente")
    motorista = st.text_input("Motorista")
    data_saida = st.date_input("Data de Saída", datetime.today())

    # Botão para adicionar dados
    if st.button("Adicionar Transferência"):
        # Verificação se todos os campos foram preenchidos
        if numero_transferencia and cidade_origem and cidade_destino and conferente and motorista and (transbordo == "Não" or (transbordo == "Sim" and cidade_transbordo and destino_final)):
            # Adicionar nova linha ao DataFrame
            new_row = pd.DataFrame([{
                "Número Transferência": numero_transferencia,
                "Cidade Origem": cidade_origem,
                "Cidade Destino": cidade_destino,
                "Quantidade Volumes": quantidade_volumes,
                "Conferente": conferente,
                "Motorista": motorista,
                "Data Saída": data_saida,
                "Cidade Transbordo": cidade_transbordo,
                "Destino Final": destino_final
            }])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # Salvar os dados no arquivo CSV
            save_data(df)
            
            st.success("Transferência adicionada com sucesso!")
        else:
            st.error("Por favor, preencha todos os campos corretamente.")

    # Exibição dos dados armazenados
    st.header("Transferências Registradas")
    st.dataframe(df)

    # Botão para baixar o arquivo CSV
    st.header("Baixar Arquivo de Transferências")
    if not df.empty:
        excel_data = to_excel(df, cidade_origem)
        if st.download_button(
            label="Baixar CSV para Impressão",
            data=excel_data,
            file_name="romaneio_cargas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            # Limpar o DataFrame e salvar o arquivo vazio
            df = pd.DataFrame(columns=["Número Transferência", "Cidade Origem", "Cidade Destino", "Quantidade Volumes", "Conferente", "Motorista", "Data Saída", "Cidade Transbordo", "Destino Final"])
            save_data(df)
            st.success("O romaneio foi baixado e a planilha foi resetada.")

if __name__ == "__main__":
    main()

