# -*- coding: utf-8 -*-


import os, os.path
import win32com.client
import pythoncom


def run_macro(file_chemin, file_name, module_name, macro_name):
    
    if os.path.exists(file_chemin):

        try:
            #xl = win32com.client.GetActiveObject("Excel.Application")
            xl = win32com.client.Dispatch("Excel.Application")
            wb_names = [wb.Name for wb in xl.Workbooks]
            if not file_name in wb_names:
                wb = xl.Workbooks.Open(os.path.abspath(file_name), ReadOnly=0)
            else:
                wb = xl.Workbooks(file_chemin)
        except:
            xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
            wb = xl.Workbooks.Open(os.path.abspath(file_chemin), ReadOnly=0)
            
            
        xl.Application.Run( file_name + "!" + module_name +"." + macro_name)
        xl.DisplayAlerts = False
        wb.Save()
        xl.Application.Quit()
        del xl
        
        message="ok - seems to work!"
    else:
        
        message = "ko - file not found"
    return message





def run_macro_from_formulaire(nom_module, nom_macro):
    
    try:
        file_name = "formulaire.xlsm"
        current_dir = os.getcwd() + r'\extractions_caceis\functions\vba'
        file_chemin = os.path.join(current_dir, file_name)
        message = run_macro(file_chemin, file_name,  nom_module, nom_macro)
        
        return message
    
    except:
        
        return "ko - sure the macro exists?"








def run_vba_function(file_chemin, file_name, module_name, macro_name):
    
    pythoncom.CoInitialize()

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(file_chemin)
    except:
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.abspath(file_chemin), ReadOnly=0)
    
    result = excel.Application.Run(file_name + "!" + module_name +"." + macro_name)
    wb.Save()
    excel.Application.Quit()
    
    return result


def run_test():
    file_name = "macros_test.xlsm"
    current_dir = os.getcwd() + r'\functions\macros'
    # current_dir = os.getcwd() + r'\macros'
    file_chemin = os.path.join(current_dir, file_name)
    module_name = "test_module"
    macro_name = "test"
    result = run_vba_function(file_chemin, file_name, module_name, macro_name)
    return result

# run_test()