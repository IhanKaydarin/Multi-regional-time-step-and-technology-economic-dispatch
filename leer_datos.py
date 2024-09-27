#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Sep 18 16:15:16 2023

@author: seck

Para agregar librerias
conda create -n spyder-env -y
conda activate spyder-env
conda install spyder-kernels openpyxl -y
o
conda install -c conda-forge spyder-kernels openpyxl -y      

"""

# Librerias
import numpy as np
import pandas as pd

def leer_datos():
    # Leer archivo excel Costos
    df1 = pd.read_excel("Costs.xlsx")
    # Seleccionar las columnas de interés
    data1 = df1[['NI','NG','NC','NT','Cost']]
    
    # Leer archivo excel Costos
    df2 = pd.read_excel("Demand.xlsx")
    # Seleccionar las columnas de interés
    data2 = df2[['NC','NT','Demand']]
    
    # Leer archivo excel Capacidad
    df3 = pd.read_excel("Capacity.xlsx")
    # Seleccionar las columnas de interés
    data3 = df3[['NI','NG','NT','MPC']]
    
    # Leer archivo excel Capacidad
    df4 = pd.read_excel("Transmission.xlsx")
    # Seleccionar las columnas de interés
    data4 = df4[['NG','NC','NT','Transmission']]
    
    # Leer archivo excel Demanda
    # df_aux = pd.read_excel("Demanda.xlsx", header=None)
    # Transponer renglones a columnas
    # df2 = df_aux.T
    # Poner primera columna como encabezados
    # df2.columns = df2.iloc[0]
    # Remover primer renglón que ya quedó como encabezado
    # df2 = df2[1:]
    # Seleccionar las columnas de interés
    # data2 = df2[['NC','NT','Demanda']]
    # Leer archivo excel Capacidad
    # df_aux = pd.read_excel("Capacidad.xlsx", header=None)
    # Transponer renglones a columnas
    # df3 = df_aux.T
    # Poner primera columna como encabezados
    # df3.columns = df3.iloc[0]
    # Remover primer renglón que ya quedó como encabezado
    # df3 = df3[1:]
    # Seleccionar las columnas de interés
    # data3 = df3[['NI','NG','NT','MPC']]
    # Leer archivo excel Transmision
    # df_aux = pd.read_excel("Transmision.xlsx", header=None)
    # Transponer renglones a columnas
    # df4 = df_aux.T
    # Poner primera columna como encabezados
    # df4.columns = df4.iloc[0]
    # Remover primer renglón que ya quedó como encabezado
    # df4 = df4[1:]
    # Seleccionar las columnas de interés
    # data4 = df4[['NG','NC','NT','Transmision']]
    
    # Tomar dataframes y pasarlos a arrays de numpy 
    array1=data1.to_numpy()
    array2=data2.to_numpy()
    array3=data3.to_numpy()
    array4=data4.to_numpy()
    
    # Salvar arreglo como archivo dat
    # Formatos para salvar datos
    formato1 = '%-6d %-6d %-6d %-6d %-3.2f'
    formato2 = '%-6d %-6d %-3.2f'
    formato3 = '%-6d %-6d %-6d %-3.2f'
    # Textos entre cada parte de datos
    parte_1 = np.array(['param N := 13;','param M := 9;','param O := 9;','param P := 24;','','','param Cost :=',''])
    parte_2 = np.array(['',';','','param Demand :=',''])
    parte_3 = np.array(['',';','','param MaximumPowerCapacity :=',''])
    parte_4 = np.array(['',';','','param Transmition_Capacity :=',''])
    parte_5 = np.array(['',';',''])
    with open('Data.dat','w') as f:
        np.savetxt(f,parte_1, delimiter=" ",fmt="%s")
        np.savetxt(f,array1,fmt=formato1)
        np.savetxt(f,parte_2, delimiter=" ",fmt="%s")
        np.savetxt(f,array2,fmt=formato2)
        np.savetxt(f,parte_3, delimiter=" ",fmt="%s")
        np.savetxt(f,array3,fmt=formato3)
        np.savetxt(f,parte_4, delimiter=" ",fmt="%s")
        np.savetxt(f,array4,fmt=formato3)
        np.savetxt(f,parte_5, delimiter=" ",fmt="%s")

