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

def generarEscenarios():
    ar1 = pd.ExcelFile("IradEnero_4.xlsx")
    df1 = ar1.parse("IradEnero_4")
    
    ar2 = pd.ExcelFile("despacho.xlsx")
    
    a = int(input("¿Desea trabajar con el despacho horario o cada 15 minutos?\n1. Horario\n2. 15 minutos\n\n"))
    
    mm = df1.columns[-1]  # Variable que contiene "Irrad PU"
    cont = 0  # Contador de pasos de tiempo
    datos = []
    
    if a == 1:
        df2 = ar2.parse("GENERADORES60")
        anadir1 = np.random.uniform(0.2,0.3,4) # Nubes entre 11am y 2pm
        anadir2 = np.random.uniform(0.15,0.25,3) # Nubes entre 2pm - 4pm
        anadir3 = np.hstack((np.random.uniform(0.65,0.75,1),np.random.uniform(0.25,0.35,2))) # Nubes entre 9am - 11am
        for i in df1[mm]:
            datos.append(df1[mm][cont])
            cont += 12
            if cont > 276:  # 276 es el límite de datos para extraer 1 irradiancia por hora
                break
        escenario1 = np.ones(int(len(df2)/4))
        escenario2 = np.ones(int(len(df2)/4))
        escenario2[10:14] = anadir1
        escenario3 = np.ones(int(len(df2)/4))
        escenario3[13:16] = anadir2
        escenario4 = np.ones(int(len(df2)/4))
        escenario4[8:11] = anadir3
        escenarios = [escenario1, escenario2, escenario3, escenario4]
    else:
        df2 = ar2.parse("GENERADORES15")
        anadir1 = np.random.uniform(0.2,0.3,12) # Nubes entre 11:30am y 2:15pm
        anadir2 = np.random.uniform(0.15,0.25,11) # Nubes entre 2:30pm - 4:45pm
        anadir3 = np.hstack((np.random.uniform(0.65,0.75,3),np.random.uniform(0.25,0.35,6))) # Nubes entre 9:30am - 11:15am
        
        for i in df1[mm]:
            datos.append(df1[mm][cont])
            cont += 3
            if cont > 285:  # 285 es el límite de datos para extraer 4 irradiancias por hora
                break

        escenario1 = np.ones(int(len(df2)/4))
        escenario2 = np.ones(int(len(df2)/4))
        escenario2[45:57] = anadir1
        escenario3 = np.ones(int(len(df2)/4))
        escenario3[56:67] = anadir2
        escenario4 = np.ones(int(len(df2)/4))
        escenario4[36:45] = anadir3
        escenarios = [escenario1, escenario2, escenario3, escenario4]
        
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
            datos.append(j*df1[cont]*df2[cont])
        matriz.append(datos)
        
    matriz = np.matrix(matriz).T
    column_names = ["Escenario 1", "Escenario 2", "Escenario 3", "Escenario 4"]
    
    dff = pd.DataFrame(matriz, columns = column_names)
    
    return dff, a