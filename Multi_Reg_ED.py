
"""
Created on Apr 26 18:50:12 2023

@author: Ulises Hurtado
"""

from pyomo.environ import *
from leer_datos import *

#Función que genera el archivo Datos.dat
leer_datos()

model = AbstractModel()



model.N = Param( within=(PositiveIntegers) )
model.M = Param( within=(PositiveIntegers) )
model.O = Param( within=(PositiveIntegers) )
model.P = Param( within=(PositiveIntegers) )



model.Technologies = RangeSet(model.N)
model.Generation_Region = RangeSet(model.M)
model.Consumption_Region = RangeSet(model.O)
model.Time_Steps = RangeSet(model.P)



model.Cost = Param (model.Technologies, model.Generation_Region, model.Consumption_Region, model.Time_Steps)
model.Demand = Param (model.Consumption_Region, model.Time_Steps)
model.MaximumPowerCapacity = Param(model.Technologies, model.Generation_Region, model.Time_Steps)
model.Transmition_Capacity = Param(model.Generation_Region, model.Consumption_Region, model.Time_Steps)



model.Energy = Var(model.Technologies, model.Generation_Region, model.Consumption_Region, model.Time_Steps, within=NonNegativeReals)



def obj_rule(model):
    return sum(sum(sum(sum(model.Cost[n,m,o,p]*model.Energy[n,m,o,p] for n in model.Technologies)
                   for m in model.Generation_Region) for o in model.Consumption_Region) for p in model.Time_Steps)
model.obj = Objective( rule=obj_rule )


def demand_match(model,o,p):
    return sum(sum(model.Energy[n,m,o,p] for n in model.Technologies)
               for m in model.Generation_Region) == model.Demand[o,p] 
model.demand_match = Constraint(model.Consumption_Region, model.Time_Steps, rule=demand_match )


def MPC(model,n,m,p):
    return sum(model.Energy[n,m,o,p] for o in model.Consumption_Region) <= model.MaximumPowerCapacity[n,m,p]
model.MPC = Constraint( model.Technologies, model.Generation_Region, model.Time_Steps, rule=MPC )


def Transmition_Capacity(model,m,o,p):
    return sum(model.Energy[n,m,o,p] for n in model.Technologies) <= model.Transmition_Capacity[m,o,p] 
model.Trasmition_Capacity = Constraint( model.Generation_Region, model.Consumption_Region, model.Time_Steps, rule=Transmition_Capacity )



instance = model.create_instance('Data.dat')
solver = SolverFactory('cplex')
#solver = SolverFactory('glpk')
solver.solve(instance)
instance.pprint()

instance.Cost.pprint()
print(type(instance.Energy))
print(value(instance.obj))

# Salida para excel
# Pasar primero a diccionario
Energia={}
for e in instance.Energy:
    Energia[e]=(value(instance.Energy[e]))
# De diccionario a dataframe, poniendo el header de la columna como energía y reseteando indices por renglon para que aparezca de 0 en adelante
df_Energia=pd.DataFrame.from_dict(Energia,orient='index',columns=['Energy']).reset_index()
# Renombra columna 'index' por 'Data' para claridad
df_Energia.rename(columns = {'index':'Data'}, inplace = True)
# La columna Data es de tipo tuple, pasarla a 4 columnas
df_Energia[['I','G','C','T']]=pd.DataFrame(df_Energia['Data'].tolist(), index=df_Energia.index)
# Reacomodar columnas
df_Energia=df_Energia[['I','G','C','T','Energy','Data']]
# Borrar columna Data
df_Energia=df_Energia.drop(['Data'], axis=1)
# De dataframe a excel
df_Energia.to_excel("Output.xlsx", index = False) 
# Imprimir primeros valores y valores distintos de cero en Energy solo para corroborar
print(df_Energia.head())
print(df_Energia[df_Energia.Energy > 0])

#####################################

#Código Dr. Joselito Medina Marín para acomodar salida a archivos Excel
# Leer archivo excel Salida
D = pd.read_excel("Output.xlsx")
# Obtención de renglones únicos y ponerlos en otro dataframe
DA = np.array(D.iloc[:,:-2])
DAU = np.unique(DA, axis=0)
DFU = pd.DataFrame(DAU, columns=D.columns[0:-2])
num_ren = DAU.shape[0]
num_col = 24
# Convertir el columna horas a una tabla
Tiempo = np.array(D.iloc[:,4])
Horas = Tiempo.reshape(num_ren,num_col)
# Encabezados para columnas de horas
enc = ['Hour ' + str(i) for i in range(1,num_col+1)]
# Converrir Horas a dataframe
DFT = pd.DataFrame(Horas, columns = enc)
# Concatenar Horas a DF con índices únicos
DFC = pd.concat([DFU, DFT], axis=1)
# Ordenarlo de acuerdo al orden I C G
DFCO = DFC.sort_values(by=["I", "C", "G"])
# Guardar archivo de Excel - Nuevo archivo
DFCO.to_excel("Results.xlsx", index = None)
# Guardar archivo de Excel - Archivo existente
with pd.ExcelWriter("TemplateResults.xlsx", mode="a",
    engine="openpyxl",
    if_sheet_exists="overlay",
) as writer:
    DFCO.iloc[:,3:].to_excel(writer, sheet_name="MATRIZ TEC RES", index = False, header=False, startrow=1, startcol=3)
    
print('End of Process')
