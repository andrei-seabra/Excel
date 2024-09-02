# Libraries used in the project
from frontend import *
from backend import *

# Control variable
option = ""

# Runs the program
while option != "6":
    # Updates the data frame
    file = read_excel(filePath, engine="openpyxl")

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

        updateClientInformation(searchClient(input("    Cliente: ")))
    elif option == "5":
        clean()
        
        print("    ---- Remover cliente ----\n")
        removeClient(searchClient(input("    Cliente: ")))

    elif option == "6":
        clean()

        print("    Saindo e salvando...")
        sleep(2)

    else:
        clean()

        print("    Opção inválida.")
        sleep(2)

    saveFile()