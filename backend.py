# Library used in the project
from time import sleep

from frontend import *
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
    global file

    # Data treatment
    if information.isdigit():
        information = int(information)
    
    # Checks if the given information is available in the file
    for column in list(file.keys()):
        if not file.loc[file[column] == information].empty:
            return file.loc[file[column] == information]
    
    print("    Cliente não encontrado.")
    sleep(2)



def getClientInformation():
    # Reference
    clientInformation = {}

    # Gets the data
    for key in list(file.keys()):
        clientInformation[key] = input(f"    {key}: ")

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



def updateClientInformation(client: DataFrame, clientInformation: dict):
    """
        Updates the given client informations.
    """
    global file

    # Checks if the the given client it's on file
    if client is None or client.empty:
        return

    # Updates the information of the given client
    for i, _ in client.iterrows():
        for key, information in clientInformation.items():
            # Data treatment
            if information.isdigit():
                information = int(information)
            
            # Updates the information
            file.at[i, key] = information

    print("    Atualizando dados...")
    sleep(2)

    file.to_excel(filePath, index=False)



def removeClient(client: DataFrame):
    """
        Removes the given client of the file.
    """
    global file

    # Reference

    # Checks if the the given client it's on file
    if client is None or client.empty:
        return
    
    # Removes the client
    file = file[file["NOME"] != client["NOME"].values[0]]

    clean()

    print("    Removendo cliente...")
    sleep(2)

    file.to_excel(filePath, index=False)