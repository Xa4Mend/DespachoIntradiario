# -*- coding: utf-8 -*-
"""
Autores:
Vladimir Álvarez Gaviria, Universidad de Antioquia, Ingeniería Eléctrica
vladimir.alvarez@udea.edu.co
Andrés Felipe Cerón Muñoz, Universidad de Antioquia, Ingeniería Eléctrica
felipe.ceron@udea.edu.co
Javier Andrés Mendoza Rocha, Universidad de Antioquia, Ingeniería Eléctrica
jandres.mendoza@udea.edu.co
"""

import numpy as np
import pandas as pd

'''
El objetivo de esta función es obtener distintos escenarios, estos escenarios modelan
4 distintos tipos de días, si es soleado o si tiene nubes a distintos horarios.
Con el objetivo de obtener la variabilidad de un panel solar de forma numérica, se hizo
uso de los valores de irradiancia en p.u. capturados en un Excel que se encontró
'''

def generarEscenarios():
    ar1 = pd.ExcelFile("IradEnero_4.xlsx")
    df1 = ar1.parse("IradEnero_4")
    
    ar2 = pd.ExcelFile("despacho.xlsx")
    
    a = int(input("¿Desea trabajar con el despacho horario o cada 15 minutos?\n1. Horario\n2. 15 minutos\n\n"))
    
    mm = df1.columns[-1]  # Variable que contiene "Irrad PU"
    cont = 0  # Contador de pasos de tiempo
    datos = [] # Vector que almacenará los datos de irradiancia p.u.
    
    if a == 1: # Se eligió trabajar con despacho horario
        df2 = ar2.parse("GENERADORES60")
        anadir1 = np.random.uniform(0.02,0.08,3) # Nubes entre 1pm y 3pm
        anadir2 = np.random.uniform(0.02,0.08,3) # Nubes entre 2pm - 4pm
        anadir3 = np.hstack((np.random.uniform(0.02,0.08,1),np.random.uniform(0.08,0.10,2))) # Nubes entre 9am - 11am
        for i in df1[mm]:
            datos.append(df1[mm][cont])
            cont += 12 # Como los datos entregados son cada 5 minutos, para pasar a la siguiente hora se le suma 60/5 = 12
            if cont > 276:  # 276 es el límite de datos para extraer 1 irradiancia por hora
                break
        # Se inicializan los escenarios en 1, para posteriormente cambiarles los datos almacenados en los vectores de anadir
        escenario1 = np.ones(int(len(df2)/4))
        escenario2 = np.ones(int(len(df2)/4))
        escenario2[12:15] = anadir1
        escenario3 = np.ones(int(len(df2)/4))
        escenario3[13:16] = anadir2
        escenario4 = np.ones(int(len(df2)/4))
        escenario4[8:11] = anadir3
        escenarios = [
            escenario1, 
            escenario2 
            # escenario3, 
            # escenario4
            ]
    else: # Se eligió trabajar con despacho cada 15 minutos
        df2 = ar2.parse("GENERADORES15")
        anadir1 = np.random.uniform(0.02,0.08,8) # Nubes entre 1pm y 3pm
        anadir2 = np.random.uniform(0.02,0.08,11) # Nubes entre 2:30pm - 4:45pm
        anadir3 = np.hstack((np.random.uniform(0.02,0.08,3),np.random.uniform(0.08,0.1,6))) # Nubes entre 9:30am - 11:15am
        
        for i in df1[mm]:
            datos.append(df1[mm][cont])
            cont += 3 # Como los datos entregados son cada 5 minutos, para pasar a la siguiente hora se le suma 15/5 = 3
            if cont > 285:  # 285 es el límite de datos para extraer 4 irradiancias por hora
                break

        escenario1 = np.ones(int(len(df2)/4))
        escenario2 = np.ones(int(len(df2)/4))
        escenario2[51:59] = anadir1
        escenario3 = np.ones(int(len(df2)/4))
        escenario3[56:67] = anadir2
        escenario4 = np.ones(int(len(df2)/4))
        escenario4[36:45] = anadir3
        escenarios = [
            escenario1, 
            escenario2 
            # escenario3, 
            # escenario4
            ]
        
    datos = np.array(datos)
    df2.set_index("nombre", inplace = True)
    df2 = df2.loc["SUPERTRINA"]["maximo"]
    
    dicc = {"Irrad PU":datos}
    df1 = pd.DataFrame(dicc)
    df1 = df1["Irrad PU"]
    
    # df1 --> Irradiancia PU
    # Supertrina (Columna máximo) --> df2
    
    matriz = []
    
    for i in escenarios:
        datos = []
        for cont,j in enumerate(i):
            datos.append(j*df1[cont]*df2[cont]) # Escenarios * Irrad_PU * SUPERTRINA["maximo"]
        matriz.append(datos)
        
    matriz = np.matrix(matriz).T
    column_names = [
        "Escenario 1", 
        "Escenario 2" 
        # "Escenario 3", 
        # "Escenario 4"
        ]
    
    dff = pd.DataFrame(matriz, columns = column_names)
    
    return dff, a