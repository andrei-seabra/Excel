# Libraries used in the project
from frontend import *
from backend import *

# Control variable
option = ""

# Runs the program
while option != "6":
    # Cleans the terminal
    clean()

    print("""
    ---- Gerenciador de Excel ----
    [1] - Visualizar clientes
    [2] - Pesquisar cliente
    [3] - Adicionar cliente
    [4] - Atualizar cliente
    [5] - Remover cliente
    [6] - Sair
    """)

    option = input("    Opção desejada: ")

    if option == "1":
        clean()

        print("    ---- Clientes ----\n")
        print(file)

        input("\n    Aperte 'Enter' para sair")

    elif option == "2":
        clean()

        print("    ---- Pesquisa ----\n")

        print(f"{searchClient(input("   Cliente desejado: "))}")
        input("\n    Aperte 'Enter' para sair")

    elif option == "3":
        clean()

        print("    ---- Cadastro de cliente ----\n")

        addClient(getClientInformation())

        clean()

        print("    Salvando dados do cliente...")
        sleep(2)

    elif option == "4":
        clean()

        print("    ---- Atualização de cliente ----\n")

        updateClientInformation(input("    Cliente (nome): "), getClientInformation())

        clean()

        print("    Atualizando dados...")
        sleep(2)

    elif option == "5":
        clean()
        
        print("    ---- Remover cliente ----\n")

        removeClient(input("    Cliente (nome): "))

        clean()

        print("    Removendo cliente...")
        sleep(2)

    elif option == "6":
        clean()

        print("    Saindo e salvando...")
        sleep(2)
        clean()

    else:
        clean()

        print("    Opção inválida.")
        sleep(2)

# Saves the excel file
file.to_excel(filePath, index=False)