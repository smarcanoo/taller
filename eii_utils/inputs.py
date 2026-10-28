from colorama import Fore

def leer_flotante(mensaje):
    while True:
        try:
            valor = float(input(mensaje))
            return valor
        except ValueError:
            print(Fore.RED + "Error: Dato inválido. Por favor, ingresa un valor numérico.")