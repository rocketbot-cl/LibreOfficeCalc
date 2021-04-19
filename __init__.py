# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

   sudo pip install <package> -t .

"""
import os
import sys


base_path = tmp_global_obj["basepath"]
cur_path = base_path + "modules" + os.sep + "LibreOfficeCalc" + os.sep + "libs" + os.sep
sys.path.append(cur_path)

import pyoo

module = GetParams("module")
global desktop_calc, doc
if module == "connect":
    result = GetParams("result")
    try:
        desktop_calc = pyoo.Desktop('localhost', 2002)

        if desktop:
            SetVar(result, True)
        else:
            SetVar(result, False)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "new":
    try:
        doc = desktop_calc.create_spreadsheet()
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "open":
    path_file = GetParams("path_file")
    try:
        doc = desktop_calc.open_spreadsheet(path_file)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "write_cell":
    import string
    cell = GetParams("cell")
    value_cell = GetParams("value_cell")
    try:
        letter_index = string.ascii_lowercase.index(cell[0].lower())
        number_index = int(cell[1])
        sheet = doc.sheets[0]
        sheet[letter_index,number_index].value = cell
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e