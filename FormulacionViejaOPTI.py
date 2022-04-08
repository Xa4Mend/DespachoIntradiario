"""
Autores:
Vladimir Álvarez Gaviria, Universidad de Antioquia, Ingeniería Eléctrica
vladimir.alvarez@udea.edu.co
Andrés Felipe Cerón Muñoz, Universidad de Antioquia, Ingeniería Eléctrica
felipe.ceron@udea.edu.co
Javier Andrés Mendoza Rocha, Universidad de Antioquia, Ingeniería Eléctrica
jandres.mendoza@udea.edu.co
"""

import pyomo.environ as p
import pandas as pd
from os import getcwd
import xlwings as xw
from string import ascii_uppercase as ABC
from xlwings import constants, Range
from obtencionEscenarios import generarEscenarios


###############################################################################################################################
# Importamos las librerías necesarias para sacar la demanda de energía en MWh
from pydataxm import *                           #Se realiza la importación de las librerias necesarias para ejecutar
import datetime as dt
from pydataxm.pydataxm import ReadDB as apiXM    #Se importa la clase que invoca el servicio
import pandas as pd
import numpy as np
###############################################################################################################################

xlCols = [i for i in ABC] + [i+j for i in ABC for j in ABC]

def obtenerDatosHoja(hojaExc):
    """
    Esta función permite obtener los datos que están en una hoja
    de Excel.
    """
    salida = pd.DataFrame(hojaExc.used_range.value).dropna(how="all").dropna(axis="columns",how="all")
    salida.columns = salida.iloc[0].values
    salida.drop(index=salida.index[0],inplace=True)
    return salida

###############################################################################################################################
# Creación de los datos de la demanda para diferentes escenarios en el Excel

# Creación de instancia
objetoAPI = pydataxm.ReadDB()                    #Se almacena el servicio en el nombre objetoAPI

# Creación de los datos para sacar los datos de la demanda

year = 2022
month = 2  # Se escoge el mes de Febrero
days = 1 # Número de días que se analizarán

# DEMANDA COMERCIAL
# Curva de carga promedio GWh
# Obtención y procesamiento

df = apiXM.request_data(objetoAPI,    #Se indica el objeto que contiene el serivicio
                        "DemaCome",           #Se indica el nombre de la métrica tal como se llama en el campo metricID
                        0,                    #Campo númerico indicando el nivel de desagregación, 1 para valores por Recurso
                        dt.date(year, month, 1),  #Corresponde a la fecha inicial de la consulta
                        dt.date(year, month, days)) #Corresponde a la fecha final de la consulta

df['Date'] = pd.to_datetime(df['Date']) #establecemos la columna Date como datetime
df.index = df['Date'] #establecemos el índice como la columna Date
df = df.drop(columns=['Id', 'Values_code', 'Date']) #eliminamos columnas innecesarias
df = round(df/10**3,4) #dividimos todos los datos por 10**6 para convertir kWh a MWh

df = df.T #obtenemos la transpuesta para una mejor visualización
list_hours = list(df.index) #obtenemos la lista de horas (1-24)
df['hour'] = [value_hour[-2:] for value_hour in list_hours] #creamos la columna hora en el df solo poniendo numeros
df.index = df['hour'] #establecemos el índice como la columna hour
df = df.drop(columns=['hour']) #eliminamos la columna hour

col = df.columns # Sacamos los nombres de las columnas del DataFrame creado
datos = []  # Aquí se almacenarán los datos de la demanda horaria
dias = ["Demanda"] # Los días serán las columnas
porcentaje = float(input("Ingrese el porcentaje de demanda con el que desea trabajar: "))
porcentaje = porcentaje / 100
min15 = [str(i+1) for i in range(24 * 4)] # Las horas serán el índice

for i in range(days):
    datos.append(np.array(df[col[i]])) # Añadimos los datos del DataFrame creado previamente

    
# Ahora vamos a añadir los datos de demanda
# para cada 15 minutos
# Se usa el llenado aleatorio usando una distribución normal
# con media igual a la demanda horaria y usando 3 desviaciones estándar  

datos_finales = [] # Esta lista almacenará los datos cada 15 minutos de los 20 días
for i in datos:
    datos15min = [] # Aquí almacenaremos los datos de la demanda cada 15 minutos
    for j in i:
        datos_alea = j + 3*np.random.randn(3) # Donde 3 es la desviación estándar
        datos15min += datos_alea.tolist()
        datos15min.append(j)
    datos_finales.append(datos15min)

mat = (np.matrix(datos_finales).T) # Creamos la matriz con los datos almacenado y la transponemos para que quede en el orden correcto
dff = pd.DataFrame(mat, columns = dias, index = min15)
df.columns = dias

# Se crean valores aleatorios en por unidad para multiplicarlos por
# la demanda real y así tener despachos más variados:

# Pasando los datos al Excel

wb = xw.Book("despacho.xlsx")
hoja_out1 = wb.sheets['despacho60viejo']
generadoresBDD = wb.sheets['GENERADORES60']
rampasBDD = wb.sheets['RAMPAS60']
demandaBDD = wb.sheets["DEMANDA60"]
demandaBDD.clear_contents()
demandaBDD.range("A1").value = df * porcentaje
demandaBDD.range("A1").value = "periodo"
demandaBDD.autofit()

###############################################################################################################################

#LECTURA DE DATOS;

df_Generar_Escenario, menu = generarEscenarios()
escenarios = len(df_Generar_Escenario.columns)

demanda = obtenerDatosHoja(demandaBDD)
auxGen = obtenerDatosHoja(generadoresBDD)
rampa = obtenerDatosHoja(rampasBDD)
    
demanda["periodo"] = demanda["periodo"].astype(int)
demanda.set_index("periodo", inplace = True)

nombresGen = auxGen.nombre.unique()
auxGen["periodo"] = auxGen["periodo"].astype(int)
geners = auxGen.set_index(['nombre','periodo']) 

rampa["tml"] = rampa["tml"].astype(int)
rampa["tmfl"] = rampa["tmfl"].astype(int)
rampa.set_index("recurso", inplace = True)

nombresRamp = rampa.index.values


###############################################################################################################################
def Despacho(escenario,geners):
    """
    Este código permite modelar matemáticamente el despacho intradiario
    para una granularidad de 1 hora.
    Toma como parámetro el número de escenarios y el DataFrame que contiene la 
    información de los recursos de generación del archivo 'despacho.xlsx'.
    """
    genersC = geners.copy()
    genersC.reset_index(inplace = True)
    genersC["maximo"].loc[genersC["nombre"]=="SUPERTRINA"] = df_Generar_Escenario["Escenario %s" %escenario].values
    geners = genersC.set_index(["nombre","periodo"]).copy()
    print("Escenario %s de generación: Potencia de SUPERTRINA [MW]" %escenario)
    for i in genersC["maximo"].loc[genersC["nombre"] == "SUPERTRINA"]: print(i)
    
    
    modelo = p.ConcreteModel("DESPACHO PROGRAMADO")
    modelo.gen = p.Set(initialize=nombresGen)
    modelo.per = p.Set(initialize=demanda.index.values)
    modelo.g = p.Var(modelo.gen,modelo.per,domain=p.PositiveReals)
    modelo.r = p.Var(modelo.per,domain=p.PositiveReals)
    modelo.u = p.Var(modelo.gen,modelo.per,domain=p.Binary)
    modelo.a = p.Var(nombresRamp,modelo.per,domain=p.Binary)
    modelo.p = p.Var(nombresRamp,modelo.per,domain=p.Binary)
    
    fun_obj1 = []
    fun_obj2 = []
    fun_obj3 = []
    
    for i in modelo.gen:
        for t in modelo.per:
            fun_obj1.append(geners.precio[i,t]*modelo.g[i,t])
            if i in nombresRamp:
                fun_obj2.append(rampa.costoarranque[i]*modelo.a[i,t])
    
    for i in modelo.per:
        fun_obj3.append(2e6*modelo.r[i])
    
    fun_obj = sum(fun_obj1) + sum(fun_obj2) + sum(fun_obj3)
    
    modelo.Obj = p.Objective(expr=fun_obj,sense=p.minimize)
    
    def R1(modelo,i,t):
        ecuacion = modelo.g[i,t] - geners.maximo[i,t]*modelo.u[i,t]
        return ecuacion <= 0
    
    modelo.R1 = p.Constraint(modelo.gen,modelo.per,rule=R1)
    
    def R2(modelo,i,t):
        ecuacion = modelo.g[i,t] - geners.minimo[i,t]*modelo.u[i,t]
        return ecuacion >= 0
    
    modelo.R2 = p.Constraint(modelo.gen,modelo.per,rule=R2)
    
    def R4(modelo, i, t):
        ci = 1 if geners.minimo[(i,1)] != 0 else 0
        u_ant = ci if t==1 else modelo.u[i,t-1]
        expr1 = modelo.a[i,t]-modelo.p[i,t]
        expr2 = modelo.u[i,t]-u_ant
        return expr1 >= expr2
    
    modelo.R4 = p.Constraint(nombresRamp,modelo.per,rule=R4)
    
    def R5(modelo,i,t):
        ecuacion = modelo.a[i,t] + modelo.p[i,t]
        return ecuacion <= 1
    
    modelo.R5 = p.Constraint(nombresRamp,modelo.per,rule=R5)
    
    #Restricción 6
    def R6(modelo, i, t):
        ci = geners.minimo[(i,1)]
        g_ant = ci if t==1 else modelo.g[i,t-1]
        expr = modelo.g[i,t]-g_ant
        return expr <= rampa["ur"][i]
    
    modelo.R6 = p.Constraint(nombresRamp, modelo.per, rule=R6)
    
    #Restricción 7
    def R7(modelo, i, t):
        ci = geners.minimo[(i,1)]
        g_ant = ci if t==1 else modelo.g[i,t-1]
        expr = g_ant-modelo.g[i,t]
        return expr <= rampa["dr"][i]
    
    modelo.R7 = p.Constraint(nombresRamp, modelo.per, rule=R7)
    
    #Restricción 8
    def R8(modelo, i, t):
        if t < rampa.tml[i]:
            return p.Constraint.Skip
        expr1 = sum(modelo.a[i,k] for k in range(modelo.per[-1]-rampa.loc[i]["tml"]+1, modelo.per[-1]))
        expr2 = modelo.u[i,t]
        return expr1 <= expr2
    
    modelo.R8 = p.Constraint(nombresRamp,modelo.per, rule=R8)
    
    #Restricción 9
    def R9(modelo, i, t):
        if t < rampa.tmfl[i]:
            return p.Constraint.Skip
        expr1 = sum(modelo.p[i,k] for k in range(modelo.per[-1]-rampa.loc[i]["tmfl"]+1, modelo.per[-1]))
        expr2 = modelo.u[i,t]
        return expr1 <= 1-expr2
    
    modelo.R9 = p.Constraint(nombresRamp,modelo.per, rule=R9)
    
    #Restricción 10
    def R10(modelo, i):
        ecuacion = []
        for t in modelo.per:
            ecuacion.append(modelo.a[i,t])
        return sum(ecuacion) <= 1
    
    modelo.R10 = p.Constraint(nombresRamp, rule=R10)
    
    #Restricción 11
    def R11(modelo,t):
        ecuacion = sum([modelo.g[i,t] for i in nombresGen])
        return ecuacion + modelo.r[t] == demanda["Demanda"][t]
    
    modelo.R11 = p.Constraint(modelo.per,rule=R11)
        
    #Definir Optimizador
    opt = p.SolverFactory('cbc')
    
    #Escribir archivo .lp
    modelo.write("archivo6.lp",io_options={"symbolic_solver_labels":True})
    
    #Ejecutar el modelo
    results = opt.solve(modelo,tee=0,logfile ="archivo6.log", keepfiles= 0,symbolic_solver_labels=True)
    
    if (results.solver.status == p.SolverStatus.ok) and (results.solver.termination_condition == p.TerminationCondition.optimal):
    
        #Imprimir Resultados
        print("\nValor óptimo de la función objetivo en Escenario %s:\n\n" %escenario, modelo.Obj())
    
    
        columnas = ["Escenario","GENERADOR"]
        for t in modelo.per:
            columnas.append(t)
    
        
        out_ = pd.DataFrame(columns=columnas)
        
        fila = 0
        
        if escenario != 1:
            aux = pd.DataFrame(hoja_out1.used_range.value).dropna(how="all").dropna(axis="columns",how="all")
            aux.index += 1
            aux = aux.index[-1] + 1
            fila += aux - 1
        
        for i in modelo.gen:
            fila += 1
            salida = []
            salida += [escenario,i]
            for t in modelo.per:
                salida.append(modelo.g[i,t].value) 
            out_.loc[fila] = salida

        out_.loc[fila+1] = [escenario,"RACIONAMIENTO"]+[modelo.r[i]() for i in modelo.per]
        out_.loc[fila+2] = [escenario,"TOTAL"] + [out_[i].loc[out_["GENERADOR"].isin(nombresGen)].sum() for i in modelo.per]
        out_.loc[fila+3] = [escenario,"DEMANDA"] + [demanda["Demanda"][i] for i in modelo.per]
        out_.loc[fila+4] = [escenario,"BALANCE"] + [out_.loc[fila+2].values.tolist()[2:][i] - demanda["Demanda"].values.tolist()[i] for i in range(0,demanda["Demanda"].index[-1])]

        if escenario == 1:

            hoja_out1.range('A1').options(index=False).value = out_

        else:

            hoja_out1.range('A%s' %aux).options(index=False).value = out_


        hoja_out1.range("{0}:{0}".format(fila+2)).color = (200,176,176)
        hoja_out1.range("{0}:{0}".format(fila+3)).color = (148,150,255)
        hoja_out1.range("{0}:{0}".format(fila+4)).color = (255,150,255)
        hoja_out1.range("{0}:{0}".format(fila+5)).color = (50,255,90)

    
    elif (results.solver.termination_condition == p.TerminationCondition.infeasible):
        print()
        print("EL PROBLEMA ES INFACTIBLE")
    
    elif(results.solver.termination_condition == p.TerminationCondition.unbounded):
        print()
        print("EL PROBLEMA ES INFACTIBLE")
    else:
        print("TERMINÓ EJECUCIÓN CON ERRORES")
        
if len(hoja_out1.used_range.address.split(":")) == 2:
    celda1,celda2 = hoja_out1.used_range.address.split(":")
    hoja_out1.range(":".join([celda1.split("$")[-1],celda2.split("$")[-1]])).api.Delete()


for i in range(1, escenarios + 1):
    Despacho(i,menu,geners)


aux = pd.DataFrame(hoja_out1.used_range.value).dropna(how="all").dropna(axis="columns",how="all")

aux.index += 1
ultimaColumna = aux.columns[-1]
aux = aux.index[-1]

hoja_out1.tables.add(hoja_out1.range("A1:{0}{1}".format(xlCols[ultimaColumna],aux)))
Range(hoja_out1.used_range.address).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
hoja_out1.autofit()
