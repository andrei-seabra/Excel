# Library used in the project
from pandas import *

# References
filePath = "Assets/Data/Template.xlsx"

# Opens the excel file
file = read_excel(filePath, engine="openpyxl")

# Methods

def searchClient(information: str):
    """
        Searchs a client by the given information (NOME, CPF, CRLV, CEP, E-MAIL, CELULAR).
    """

    # Data treatment
    if information.isdigit():
        information = int(information)
    
    # Checks if the given information is available in the file
    for column in list(file.keys()):
        if not file.loc[file[column] == information].empty:
            return file.loc[file[column] == information]



def getClientInformation():
    # References
    keys = ["NOME", "CPF", "CRLV"]
    clientInformation = {}

    # Gets the data
    clientInformation["NOME"] = input("    Nome: ")
    clientInformation["CPF"] = input("    CPF: ")
    clientInformation["CRLV"] = input("    CRLV: ")
    clientInformation["CEP"] = input("    CEP: ")
    clientInformation["E-MAIL"] = input("    E-MAIL: ")
    clientInformation["CELULAR"] = input("    CELULAR: ")

    return clientInformation



def addClient(clientInformation: dict):
    """
        Adds a new client to the file with the given information.
    """
    global file

    # Organizes the information into excel format
    newClient = DataFrame([clientInformation])
    file = concat([file, newClient], ignore_index=True) # Adds the new client data to the file

    # Saves the information of the new client
    file.to_excel(filePath, index=False)



def updateClientInformation(client: str, clientInformation: dict):
    """
        Updates the given client informations.
    """
    global file

    # Reference
    row = file[file["NOME"] == client]

    # Checks if the the given client it's on file
    if row.empty:
        print("    Cliente não encontrado.")
        return

    # Updates the information of the given client
    for i, _ in row.iterrows():
        for key, information in clientInformation.items():
            # Data treatment
            if information.isdigit():
                information = int(information)
            
            # Updates the information
            file.at[i, key] = information

    file.to_excel(filePath, index=False)



def removeClient(client: str):
    """
        Removes the given client of the file.
    """
    global file

    # Reference
    row = file[file["NOME"] == client]

    # Checks if the the given client it's on file
    if row.empty:
        return
    
    # Removes the client
    file = file[file["NOME"] != client]