import os
import tkinter as tk

def organizar(path):
    """
    Funcion para organizar los archivos en las carpetas pertinentes. La asignacion se basa en el nombre de cada archivo
    :param path: El directorio donde se encuentran todos los archivos sin clasificar antes del analisis
    :return:
    """
    try:
        os.mkdir(f'{path}Audiencias')
    except:
        print("Audiencias ya existe")
    try:
        os.mkdir(f'{path}Ordenes')
    except:
        print("Ordenes ya existe")
    try:
        os.mkdir(f'{path}Cateos')
    except:
        print("Cateos ya existe")
    try:
        os.mkdir(f'{path}Medidas')
    except:
        print("Medidas de protección ya existe")
    try:
        os.mkdir(f'{path}Informe diario')
    except:
        print("Informe diario ya existe")
    try:
        os.mkdir(f'{path}Otras')
    except:
        print("Otras ya existe")
    file_list = os.listdir(path)
    file_list = [file for file in file_list if os.path.isfile(f"{path}{file}")]
    for file in file_list:
        if "xlsx" not in file.lower():
            os.rename(f'{path}{file}', f'{path}Otras/{file}')
        elif "audiencia" in file.lower() or "formato informe concentrado" in file.lower():
            os.rename(f'{path}{file}', f'{path}Audiencias/{file}')
        elif "orden" in file.lower():
            os.rename(f'{path}{file}', f'{path}Ordenes/{file}')
        elif "cateo" in file.lower():
            os.rename(f'{path}{file}', f'{path}Cateos/{file}')
        elif "medidas" in file.lower():
            os.rename(f'{path}{file}', f'{path}Medidas/{file}')
        elif "ejecutivo" in file.lower():
            os.rename(f'{path}{file}', f'{path}Informe diario/{file}')
        else:
            os.rename(f'{path}{file}', f'{path}Otras/{file}')