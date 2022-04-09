"""
Autores:
Vladimir Álvarez Gaviria, Universidad de Antioquia, Ingeniería Eléctrica
vladimir.alvarez@udea.edu.co
Andrés Felipe Cerón Muñoz, Universidad de Antioquia, Ingeniería Eléctrica
felipe.ceron@udea.edu.co
Javier Andrés Mendoza Rocha, Universidad de Antioquia, Ingeniería Eléctrica
jandres.mendoza@udea.edu.co
"""

import pyomo.environ as p  # Librería de optimización
import pandas as pd
import xlwings as xw  # Librería para modificar el archivo de Excel mientras está abierto
from string import ascii_uppercase as ABC  # Librería para crear los nombres de las columnas de Excel
from xlwings import constants, Range  # Parámetros para usar el centrado de las celdas de Excel 
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
dff = pd.DataFrame(mat, columns = dias, index = min15) # Se almacenan los datos sacados de XM de la demanda cada 15 minutos
df.columns = dias # Se almacenan los datos sacados de XM de la demanda horaria

# Se crean valores aleatorios en por unidad para multiplicarlos por
# la demanda real y así tener despachos más variados:

# Pasando los datos al Excel

wb = xw.Book("despacho.xlsx")

# Lectura del almacenamiento de datos y visualización de resultados de despacho para un período de 15 minutos
hoja_out2 = wb.sheets['despacho15']
generadoresBD = wb.sheets['GENERADORES15']
rampasBD = wb.sheets['RAMPAS15']
demandaBD = wb.sheets["DEMANDA15"]
demandaBD.clear_contents()
demandaBD.range("A1").value = dff * porcentaje
demandaBD.range("A1").value = "periodo"
demandaBD.autofit()

# Lectura del almacenamiento de datos y visualización de resultados de despacho para un período de 60 minutos
hoja_out1 = wb.sheets['despacho60']
generadoresBDD = wb.sheets['GENERADORES60']
rampasBDD = wb.sheets['RAMPAS60']
demandaBDD = wb.sheets["DEMANDA60"]
demandaBDD.clear_contents()
demandaBDD.range("A1").value = df * porcentaje
demandaBDD.range("A1").value = "periodo"
demandaBDD.autofit()

###############################################################################################################################

#LECTURA DE DATOS;

df_Generar_Escenario, menu = generarEscenarios() # La variable menú es la que me indica si estamos trabajando el despacho
                                                 # horario o cada 15 minutos
escenarios = len(df_Generar_Escenario.columns)

if menu == 1:
    demanda = obtenerDatosHoja(demandaBDD)
    auxGen = obtenerDatosHoja(generadoresBDD)
    rampa = obtenerDatosHoja(rampasBDD)
    
else:
    demanda = obtenerDatosHoja(demandaBD)
    auxGen = obtenerDatosHoja(generadoresBD)
    rampa = obtenerDatosHoja(rampasBD)

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
def Despacho(escenario,menu,geners): # menu = 1 --> Toma los datos horario, else: Toma los datos de 15 minutos
    """
    Este código permite modelar matemáticamente el despacho intradiario
    para una granularidad de 60 y 15 minutos.
    Toma como parámetro el número de escenarios, la opción elegida en
    el menú y el DataFrame que contiene la información de los
    recursos de generación del archivo 'despacho.xlsx'.
    """
    #------------------------------------------------------------------------------------------------
    # Cargado de datos de escenario de generación para el recurso solar en función del escenario
    #------------------------------------------------------------------------------------------------
    genersC = geners.copy()
    genersC.reset_index(inplace = True)
    genersC["maximo"].loc[genersC["nombre"]=="SUPERTRINA"] = df_Generar_Escenario["Escenario %s" %escenario].values
    geners = genersC.set_index(["nombre","periodo"]).copy()
    print("Escenario %s de generación: Potencia de SUPERTRINA [MW]" %escenario)
    for i in genersC["maximo"].loc[genersC["nombre"] == "SUPERTRINA"]: print(i)
    
    #------------------------------------------------------------------------------------------------
    # Inicialización de la modelación matemática y declaración de variables matemáticas
    #------------------------------------------------------------------------------------------------
    
    modelo = p.ConcreteModel("DESPACHO PROGRAMADO")
    modelo.gen = p.Set(initialize=nombresGen)
    modelo.per = p.Set(initialize=demanda.index.values)
    modelo.g = p.Var(modelo.gen,modelo.per,domain=p.PositiveReals)
    modelo.r = p.Var(modelo.per,domain=p.PositiveReals)
    modelo.u = p.Var(modelo.gen,modelo.per,domain=p.Binary)
    modelo.a = p.Var(nombresRamp,modelo.per,domain=p.Binary)
    modelo.p = p.Var(nombresRamp,modelo.per,domain=p.Binary)
    modelo.subper = p.RangeSet(1,60) if menu == 1 else p.RangeSet(1,15)
    modelo.pg = p.Var(nombresRamp,modelo.per,modelo.subper,domain=p.PositiveReals)
    
    #------------------------------------------------------------------------------------------------
    # Implementación de la función objetivo
    #------------------------------------------------------------------------------------------------
    
    fun_obj1 = []  # Término del precio de oferta del generador por su generación
    fun_obj2 = []  # Término del costo de arranque del recurso térmico
    fun_obj3 = []  # Término del racionamiento
    
    for i in modelo.gen:
        for t in modelo.per:
            fun_obj1.append(geners.precio[i,t]*modelo.g[i,t])
            if i in nombresRamp:
                fun_obj2.append(rampa.costoarranque[i]*modelo.a[i,t])
    
    for i in modelo.per:
        fun_obj3.append(2e6*modelo.r[i])
    
    fun_obj = sum(fun_obj1) + sum(fun_obj2) + sum(fun_obj3)
    
    modelo.Obj = p.Objective(expr=fun_obj,sense=p.minimize)
    #------------------------------------------------------------------------------------------------
    
    #------------------------------------------------------------------------------------------------
    # Implementación de restricciones
    #------------------------------------------------------------------------------------------------
    
    def R1(modelo,i,t):
        """
        Restricción de la ecuación (2) del artículo
        """
        ecuacion = modelo.g[i,t] - geners.maximo[i,t]*modelo.u[i,t]
        return ecuacion <= 0
    
    modelo.R1 = p.Constraint(modelo.gen,modelo.per,rule=R1)
    
    def R2(modelo,i,t):
        """
        Restricción de la ecuación (3) del artículo
        """
        ecuacion = modelo.g[i,t] - geners.minimo[i,t]*modelo.u[i,t]
        return ecuacion >= 0
    
    modelo.R2 = p.Constraint(modelo.gen,modelo.per,rule=R2)
    
    def R4(modelo, i, t):
        """
        Restricción de la ecuación (4) del artículo
        """
        ci = 1 if geners.minimo[(i,1)] != 0 else 0
        u_ant = ci if t==1 else modelo.u[i,t-1]
        expr1 = modelo.a[i,t]-modelo.p[i,t]
        expr2 = modelo.u[i,t]-u_ant
        return expr1 >= expr2
    
    modelo.R4 = p.Constraint(nombresRamp,modelo.per,rule=R4)
    
    def R5(modelo,i,t):
        """
        Restricción de la ecuación (5) del artículo
        """
        ecuacion = modelo.a[i,t] + modelo.p[i,t]
        return ecuacion <= 1
    
    modelo.R5 = p.Constraint(nombresRamp,modelo.per,rule=R5)
    
    #Restricción 6
    def R6(modelo, i, t, k):
        """
        Restricción de la ecuación (8) del artículo
        """
        expr = modelo.pg[i,t,k] - modelo.pg[i,t,k-1]
        return expr <= rampa["vtc"][i]
    
    modelo.R6 = p.Constraint(nombresRamp, modelo.per, p.RangeSet(2,modelo.subper[-1]), rule=R6)
    
    def R6prima(modelo,i,t):
        """
        Restricción de la ecuación (9) del artículo
        """
        expr = modelo.pg[i,t,modelo.subper[1]] - modelo.pg[i,t-1,modelo.subper[-1]]
        return expr <= rampa["vtc"][i]

    modelo.R6prima = p.Constraint(nombresRamp, p.RangeSet(2,modelo.per[-1]), rule=R6prima)

    #Restricción 7
    def R7(modelo, i, t, k):
        """
        Restricción de la ecuación (10) del artículo
        """
        expr = modelo.pg[i,t,k-1] - modelo.pg[i,t,k]
        return expr <= rampa["vtd"][i]
    
    modelo.R7 = p.Constraint(nombresRamp, modelo.per, p.RangeSet(2,modelo.subper[-1]), rule=R7)
    
    def R7prima(modelo,i,t):
        """
        Restricción de la ecuación (11) del artículo
        """
        expr = modelo.pg[i,t-1,modelo.subper[-1]] - modelo.pg[i,t,modelo.subper[1]]
        return expr <= rampa["vtd"][i]

    modelo.R7prima = p.Constraint(nombresRamp, p.RangeSet(2,modelo.per[-1]), rule=R7prima)

    #Restricción 8
    def R8(modelo, i, t):
        """
        Restricción de la ecuación (17) del artículo
        """
        if t < rampa.tml[i]:
            return p.Constraint.Skip
        expr1 = sum(modelo.a[i,k] for k in range(modelo.per[-1]-rampa.loc[i]["tml"]+1, modelo.per[-1]))
        expr2 = modelo.u[i,t]
        return expr1 <= expr2
    
    modelo.R8 = p.Constraint(nombresRamp,modelo.per, rule=R8)
    
    #Restricción 9
    def R9(modelo, i, t):
        """
        Restricción de la ecuación (18) del artículo
        """
        if t < rampa.tmfl[i]:
            return p.Constraint.Skip
        expr1 = sum(modelo.p[i,k] for k in range(modelo.per[-1]-rampa.loc[i]["tmfl"]+1, modelo.per[-1]))
        expr2 = modelo.u[i,t]
        return expr1 <= 1-expr2
    
    modelo.R9 = p.Constraint(nombresRamp,modelo.per, rule=R9)
    
    #Restricción 10
    def R10(modelo, i):
        """
        Restricción de la ecuación (20)
        """
        ecuacion = []
        for t in modelo.per:
            ecuacion.append(modelo.a[i,t])
        return sum(ecuacion) <= 1
    
    modelo.R10 = p.Constraint(nombresRamp, rule=R10)
    
    #Restricción 11
    def R11(modelo,t):
        """
        Restricción de la ecuación (19) del artículo
        """
        ecuacion = sum([modelo.g[i,t] for i in nombresGen])
        return ecuacion + modelo.r[t] == demanda["Demanda"][t]
    
    modelo.R11 = p.Constraint(modelo.per,rule=R11)
    
    # Restricción 12
    def R12(modelo, i, t, k):
        """
        Restricción de la ecuación (12) del artículo
        """
        expr = modelo.pg[i,t,k]
        return expr <= geners.maximo[i,t]
    
    modelo.R12 = p.Constraint(nombresRamp, modelo.per, modelo.subper, rule=R12)
    
    # Restricción 13
    def R13(modelo, i, t, k):
        """
        Restricción de la ecuación (13) del artículo
        """
        expr = modelo.pg[i,t,k]
        return expr >= geners.minimo[i,t] * modelo.u[i,t]
    
    modelo.R13 = p.Constraint(nombresRamp, modelo.per, modelo.subper, rule=R13)
    
    # Restricción 14
    def R14(modelo, i, t):
        """
        Restricción de la ecuación (15) del artículo
        """
        estadoTmenos1 = 0 if t-1 == 0 else modelo.pg[i,t-1,modelo.subper[-1]]
        sumatorio = sum(modelo.pg[i,t,k] for k in range(1,modelo.subper[-1]))
        expr1 = estadoTmenos1 + modelo.pg[i,t,modelo.subper[-1]] + 2*sumatorio
        expr2 = 2 * modelo.subper[-1] * modelo.g[i,t]
        return expr1 == expr2
    
    modelo.R14 = p.Constraint(nombresRamp,p.RangeSet(1,modelo.per[-1]),rule=R14)

    def R15(modelo, i):
        """
        Restricción de la ecuación (16)
        """
        expr = modelo.pg[i,1,1]
        return expr <= rampa["vtc"][i]
    
    modelo.R15 = p.Constraint(nombresRamp,rule=R15)
    

    #------------------------------------------------------------------------------------------------
        
        
    # Definir Optimizador
    opt = p.SolverFactory('cbc')
    
    # Escribir archivo .lp
    modelo.write("archivo6.lp",io_options={"symbolic_solver_labels":True})
    
    # Ejecutar el modelo
    results = opt.solve(modelo,tee=0,logfile ="archivo6.log", keepfiles= 0,symbolic_solver_labels=True)
    
    if (results.solver.status == p.SolverStatus.ok) and (results.solver.termination_condition == p.TerminationCondition.optimal):
    
        # Imprimir Resultados del valor óptimo de la función objetivo en la consola en función del número 
        # final de subperiodos
        print("\nValor óptimo de la función objetivo en Escenario %s:\n\n" %escenario, modelo.subper[-1]/60 * modelo.Obj())
    
    
        columnas = ["Escenario","GENERADOR"]
        for t in modelo.per:
            columnas.append(t)
    
        
        out_ = pd.DataFrame(columns=columnas)
        
        fila = 0
        # Sección auxiliar que permite obtener el dato de la fila final de escritura del escenario
        if escenario != 1:
            if menu == 1:
                aux = pd.DataFrame(hoja_out1.used_range.value).dropna(how="all").dropna(axis="columns",how="all")
            else:
                aux = pd.DataFrame(hoja_out2.used_range.value).dropna(how="all").dropna(axis="columns",how="all")
            aux.index += 1
            aux = aux.index[-1] + 1
            fila += aux - 1
        
        # Ciclo de creación del DataFrame del despacho bajo un escenario
        for i in modelo.gen:
            fila += 1
            salida = []
            salida += [escenario,i]
            for t in modelo.per:
                salida.append(modelo.g[i,t].value) 
            out_.loc[fila] = salida

        # Filas adicionales de RACIONAMIENTO, TOTAL, DEMANDA y BALANCE
        out_.loc[fila+1] = [escenario,"RACIONAMIENTO"]+[modelo.r[i]() for i in modelo.per]
        out_.loc[fila+2] = [escenario,"TOTAL"] + [out_[i].loc[out_["GENERADOR"].isin(nombresGen)].sum() for i in modelo.per]
        out_.loc[fila+3] = [escenario,"DEMANDA"] + [demanda["Demanda"][i] for i in modelo.per]
        out_.loc[fila+4] = [escenario,"BALANCE"] + [out_.loc[fila+2].values.tolist()[2:][i] - demanda["Demanda"].values.tolist()[i] for i in range(0,demanda["Demanda"].index[-1])]

        # Definición del lugar de escritura e inserción del despacho en el Excel dependiendo del escenario

        if escenario == 1:
            if menu == 1:
                hoja_out1.range('A1').options(index=False).value = out_
            else:
                hoja_out2.range('A1').options(index=False).value = out_
        else:
            if menu == 1:
                hoja_out1.range('A%s' %aux).options(index=False).value = out_
            else:
                hoja_out2.range('A%s' %aux).options(index=False).value = out_

        # Definición de colores de las filas adicionales

        if menu == 1:
            hoja_out1.range("{0}:{0}".format(fila+2)).color = (200,176,176)
            hoja_out1.range("{0}:{0}".format(fila+3)).color = (148,150,255)
            hoja_out1.range("{0}:{0}".format(fila+4)).color = (255,150,255)
            hoja_out1.range("{0}:{0}".format(fila+5)).color = (50,255,90)
        else:
            hoja_out2.range("{0}:{0}".format(fila+2)).color = (200,176,176)
            hoja_out2.range("{0}:{0}".format(fila+3)).color = (148,150,255)
            hoja_out2.range("{0}:{0}".format(fila+4)).color = (255,150,255)
            hoja_out2.range("{0}:{0}".format(fila+5)).color = (50,255,90)
    
    elif (results.solver.termination_condition == p.TerminationCondition.infeasible):
        print()
        print("EL PROBLEMA ES INFACTIBLE")
    
    elif(results.solver.termination_condition == p.TerminationCondition.unbounded):
        print()
        print("EL PROBLEMA ES INFACTIBLE")
    else:
        print("TERMINÓ EJECUCIÓN CON ERRORES")
        

# Borrado de contenido de hojas

if menu == 1:
    if len(hoja_out1.used_range.address.split(":")) == 2:  # Condición que asegura que se borre una hoja con contenido
        celda1,celda2 = hoja_out1.used_range.address.split(":")
        hoja_out1.range(":".join([celda1.split("$")[-1],celda2.split("$")[-1]])).api.Delete()
else:
    if len(hoja_out2.used_range.address.split(":")) == 2:
        celda1,celda2 = hoja_out2.used_range.address.split(":")
        hoja_out2.range(":".join([celda1.split("$")[-1],celda2.split("$")[-1]])).api.Delete()


#------------------------------------------------------------------------------------------------
# Calculando el mejor despacho para los distintos escenarios e inserción directa al Excel
#------------------------------------------------------------------------------------------------
for i in range(1, escenarios + 1): 
    Despacho(i,menu,geners)
#------------------------------------------------------------------------------------------------

# Proceso auxiliar para determinar la columna y fila final del Excel para generar 
# los filtros de tabla

if menu == 1:
    aux = pd.DataFrame(hoja_out1.used_range.value).dropna(how="all").dropna(axis="columns",how="all")
else:
    aux = pd.DataFrame(hoja_out2.used_range.value).dropna(how="all").dropna(axis="columns",how="all")
aux.index += 1
ultimaColumna = aux.columns[-1]
aux = aux.index[-1]

#------------------------------------------------------------------------------------------------
# Inserción de la tabla con filtros al Excel de despacho
#------------------------------------------------------------------------------------------------

if menu == 1:
    hoja_out1.tables.add(hoja_out1.range("A1:{0}{1}".format(xlCols[ultimaColumna],aux)))
    Range(hoja_out1.used_range.address).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    hoja_out1.autofit()
else:
    hoja_out2.tables.add(hoja_out2.range("A1:{0}{1}".format(xlCols[ultimaColumna],aux)))
    Range(hoja_out2.used_range.address).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    hoja_out2.autofit()
#------------------------------------------------------------------------------------------------