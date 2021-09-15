# -*- coding: utf-8 -*-
"""
Created on Tue Aug 25 19:47:26 2020

@author: evule
"""

import pandas as pd


Base=pd.ExcelFile('Extracto del dia.xlsx')

BaseARS=Base.parse('285')
SaldoInicialARS=pd.DataFrame({'SaldoInicial':[BaseARS.iloc[0]['CREDITO']], 'Moneda':['ARS']})
BaseARS=BaseARS[1:-4]
BaseARS['Moneda']='ARS'

BaseUSD=Base.parse('80285')
SaldoInicialUSD=pd.DataFrame({'SaldoInicial':[BaseUSD.iloc[0]['CREDITO']], 'Moneda':['USD']})
BaseUSD=BaseUSD[1:-4]
BaseUSD['Moneda']='USD'

BaseTotal=pd.concat([BaseARS,BaseUSD])
SaldoInicial=pd.concat([SaldoInicialARS,SaldoInicialUSD])

Metricas=BaseTotal.groupby('Moneda')['DEBITO','CREDITO'].sum()
Metricas=pd.merge(Metricas,SaldoInicial, on='Moneda')
Metricas['SaldoFinal']=Metricas['SaldoInicial']+Metricas['CREDITO']-Metricas['DEBITO']
Metricas.set_index('Moneda', inplace=True)
#print(Metricas)


PrincipalesCreditos=BaseTotal.groupby(['Moneda','DESCRIPCION'])['CREDITO'].sum()
PrincipalesCreditos=PrincipalesCreditos.groupby(level=0, group_keys=False)
PrincipalesCreditos=pd.DataFrame(PrincipalesCreditos.apply(lambda x: x.sort_values(ascending=False).head(3)))
TresCreditos=pd.DataFrame(PrincipalesCreditos.groupby('Moneda').sum())
TresCreditos.reset_index(inplace=True)
TresCreditos.columns=(['Moneda','TresCreditos'])

PrincipalesDebitos=BaseTotal.groupby(['Moneda','DESCRIPCION'])['DEBITO'].sum()
PrincipalesDebitos=PrincipalesDebitos.groupby(level=0, group_keys=False)
PrincipalesDebitos=pd.DataFrame(PrincipalesDebitos.apply(lambda x: x.sort_values(ascending=False).head(3)))
TresDebitos=pd.DataFrame(PrincipalesDebitos.groupby('Moneda').sum())
TresDebitos.reset_index(inplace=True)
TresDebitos.columns=(['Moneda','TresDebitos'])


Resto=pd.merge(pd.merge(Metricas,TresCreditos, on='Moneda'), TresDebitos, on='Moneda')
Resto['RestoCreditos']=Resto['CREDITO']-Resto['TresCreditos']
Resto['RestoDebitos']=Resto['DEBITO']-Resto['TresDebitos']

RestoCreditos=pd.DataFrame({'Moneda':Resto['Moneda'], 'DESCRIPCION': ['Resto Creditos']*2, 'CREDITO':Resto['RestoCreditos']})
RestoCreditos=pd.DataFrame(RestoCreditos.groupby(['Moneda','DESCRIPCION'])['CREDITO'].sum())
PrincipalesCreditos=PrincipalesCreditos.append(RestoCreditos).sort_values(by='Moneda')

TotalCreditos=pd.DataFrame({'DESCRIPCION':['Total Creditos']*2, 'CREDITO':PrincipalesCreditos.groupby('Moneda')['CREDITO'].sum()})
idx=pd.MultiIndex.from_arrays([['ARS','USD'], TotalCreditos['DESCRIPCION']])
TotalCreditos.index=idx
TotalCreditos=TotalCreditos.drop('DESCRIPCION',1)
PrincipalesCreditos=PrincipalesCreditos.append(TotalCreditos).sort_values(by='Moneda')
#print(PrincipalesCreditos)

RestoDebitos=pd.DataFrame({'Moneda':Resto['Moneda'], 'DESCRIPCION': ['Resto Debitos']*2, 'DEBITO':Resto['RestoDebitos']})
RestoDebitos=pd.DataFrame(RestoDebitos.groupby(['Moneda','DESCRIPCION'])['DEBITO'].sum())
PrincipalesDebitos=PrincipalesDebitos.append(RestoDebitos).sort_values(by='Moneda')

TotalDebitos=pd.DataFrame({'DESCRIPCION':['Total Debitos']*2, 'DEBITO':PrincipalesDebitos.groupby('Moneda')['DEBITO'].sum()})
idx=pd.MultiIndex.from_arrays([['ARS','USD'], TotalDebitos['DESCRIPCION']])
TotalDebitos.index=idx
TotalDebitos=TotalDebitos.drop('DESCRIPCION',1)
PrincipalesDebitos=PrincipalesDebitos.append(TotalDebitos).sort_values(by='Moneda')
#print(PrincipalesDebitos)


BaseTotal['HoraEntera']=BaseTotal['HORA'].apply(lambda x: x[:2]).astype(int)+1
Cuadro=pd.DataFrame(BaseTotal.groupby(['Moneda','HoraEntera']).agg({'CREDITO':['sum','mean','max','min'],'DEBITO':['sum','mean','max','min']}))
Cuadro.columns=['Creditos','Prom Creditos','Max Credito','Min Credito','Debitos','Prom Debitos','Max Debito','Min Debito']
Cuadro['Saldos Acumulados']=Cuadro.groupby('Moneda')['Creditos'].cumsum()-Cuadro.groupby('Moneda')['Debitos'].cumsum()
Cuadro=Cuadro[['Saldos Acumulados','Creditos','Debitos','Prom Creditos','Max Credito','Min Credito','Prom Debitos','Max Debito','Min Debito']]
#print(Cuadro)

BaseTotal=BaseTotal.sort_values(by=['Moneda','HORA'])
BaseTotal['LiquidezIntradiaria']=BaseTotal.groupby(['Moneda'])['CREDITO'].cumsum()-BaseTotal.groupby(['Moneda'])['DEBITO'].cumsum()
UsoLiquidezIntradiaria=pd.DataFrame(BaseTotal.groupby('Moneda').agg({'LiquidezIntradiaria':['min','max']}))
UsoLiquidezIntradiaria.columns=['UsoLiqIntradNeg','UsoLiqIntradPos']
UsoLiquidezIntradiaria['UsoLiqIntradNeg']=UsoLiquidezIntradiaria['UsoLiqIntradNeg']*-1
#print(UsoLiquidezIntradiaria)

EgresoFondos=Cuadro.groupby('Moneda')['Debitos'].cumsum()/Metricas['DEBITO']
EgresoFondos=(EgresoFondos.unstack(level='HoraEntera')).transpose()
EgresoFondos=EgresoFondos.fillna(method='bfill')
EgresoFondos=EgresoFondos.loc[(EgresoFondos.index>=10) & (EgresoFondos.index<=20)]
#print(EgresoFondos)

#Estres:
Exclusiones=['CONCERTACION PA.PASIVOS','INTERFAZ AFIP (OSIRIS)']
BaseTotal.reset_index(inplace=True)
BaseTotal['HoraEstres1']=BaseTotal['HORA']
BaseTotal.loc[BaseTotal['DEBITO']>BaseTotal['CREDITO'],['HoraEstres1']]=((BaseTotal['HoraEstres1'].apply(lambda x: x[:2])).astype(int)+1).astype('str')+':'+BaseTotal['HoraEstres1'].apply(lambda x: x[3:])
BaseTotal.loc[BaseTotal['DESCRIPCION'].isin(Exclusiones),'DEBITO']=0


BaseTotal=BaseTotal.sort_values(by=['Moneda','HoraEstres1'])
BaseTotal['LiquidezIntradiariaEstres1']=BaseTotal.groupby(['Moneda'])['CREDITO'].cumsum()-BaseTotal.groupby(['Moneda'])['DEBITO'].cumsum()
UsoLiquidezIntradiariaEstres1=pd.DataFrame(BaseTotal.groupby('Moneda').agg({'LiquidezIntradiariaEstres1':['min','max']}))
UsoLiquidezIntradiariaEstres1.columns=['UsoLiqIntradNegEstres1','UsoLiqIntradPosEstres1']
UsoLiquidezIntradiariaEstres1['UsoLiqIntradNegEstres1']=UsoLiquidezIntradiariaEstres1['UsoLiqIntradNegEstres1']*-1
print(UsoLiquidezIntradiariaEstres1)


#print(BaseTotal)

BaseTotal.to_excel(r'C:\Users\evule\Desktop\Python\Python ALM\Salida2.xlsx', index = False)