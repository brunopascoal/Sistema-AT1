import streamlit as st
from streamlit_option_menu import option_menu

# Função para a homepage
def homepage():
    st.title("Relatórios Grupo Toniello")
    st.write("Bem-vindo à Homepage. Escolha uma Filial na aba lateral.")

# Função para carregar o script do cliente
def carregar_script(nome_cliente):
    if nome_cliente == "AT1":
        import AT1
        AT1.main()
    elif nome_cliente == "PC1":
        import PC1
        PC1.main()
        pass

with st.sidebar:
    #st.markdown("### :clipboard: Menu de Clientes")  # Altere o ícone conforme desejado

    opcao = option_menu(
        menu_title="Menu",
        options=["Homepage", "AT1", "PC1"],
        icons=["house", "star", "star"],  # Ícones para cada opção
        default_index=0  # Define a homepage como padrão
    )

# Verifica qual opção foi selecionada e carrega a página correspondente
if opcao == "Homepage":
    homepage()
else:
    carregar_script(opcao)

