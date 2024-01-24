#!/usr/bin/env python
# coding: utf-8

# In[27]:


import pandas as pd
import xlwings as xw
import os
from datetime import datetime
from pathlib import Path
from decimal import Decimal, ROUND_DOWN

# Caminho do diretório
caminho_diretorio = Path("C:\\Users\\Vertys\\Documents\\Export de dados")

#Criação de um dataframe para salvar depois com os vlaores gerados
df1=pd.DataFrame({
                    "Dia":[''],
                    "Pv1":[''],
                    "Pv2":[''],
                    "Eficiência":['']
                                })



#Lista dos valores da soma das potências diarias
potencia_total=[]

#Lista de valores 
gerado_pv1=[]
gerado_pv2=[]

#Dic para conter os dias e plotar no DF
dias={}



# Lista os arquivos no diretório
pastas = caminho_diretorio.iterdir()


#arquivos=pasta.iterdir()
for pasta in pastas:
    
    #Pegando o caminho da pasta pelo metodo Path.
    caminho=Path(pasta)
    
    #Permitindo interar sobre o camimho da pasta, selecionando assim seus arquivos
    arquivos=caminho.iterdir()
    
   
    try:
        #Intera sobre os arquivos da pasta selecionada
        for arquivo in arquivos:

            try:   
                
                #Ler Excel
                df = pd.read_excel(arquivo)
              
                #Ler pv1
                pt01=df['PV1 Power(W)']
                pt1=pt01

                #Ler pv2
                pt02=df['PV2 Power(W)']
                pt2=pt02

                #Potência do dia
                p=df['Yield(kWh)']
                pot=p.iloc[-10]

                potencia_total.append((pot))

                #Time do dia
                tempes=df['Time']  
                tempo = [str(t) for t in tempes]

                
                #Dia, para colocar no dic e atribuir ao excel
                temporal=tempo[0]

                #Potência do dia
                t=df['Time']
                tem=t.iloc[-10]

                
                #Criando valores para serem interados e usados como indices
                i=0
                t=0
                
                #Lista para add os valres das aréas de cada PV
                area1=[]
                area2=[]

                try:


                    #Intera sobre os valores de tempo para pegar o tempo exato de cada ponto de leitura
                    for te in tempo:

                        #Intera sobre i + 1
                        i+=1

                        #Pega o valor
                        tempo1=te[11:]


                        #Pega o valor seguinte da lista
                        tempo001=tempo[i]
                        tempo01=tempo001[11:]


                        #Formatar valores e ver sua difenreça de tempo
                        f = '%H:%M:%S'
                        dif = (datetime.strptime(tempo01, f) - datetime.strptime(tempo1, f)).total_seconds()

                        #Calcula a diferneça e coverte para horas.
                        total_tem=(dif/60)/60

                        #Trucando um valor de horario
                        t_truncado = Decimal(str(total_tem)).quantize(Decimal('0.00'), rounding=ROUND_DOWN)

                        #Converte o valor truncado para float
                        t_truncado_float = float(t_truncado)
                                              

                        #Divide o valor de tmepo por /2
                        tem=t_truncado_float/2

                        #Pega o valor de potência no indice do valor de t
                        v=pt1[t]


                        #Calculando area da pv1
                        valor_base1=pt1[i]
                        are1=(((v+valor_base1) * tem)/1000)


                        #Trunca os valor da area1 
                        n1_truncado = Decimal(str(are1)).quantize(Decimal('0.00'), rounding=ROUND_DOWN)


                        #Converte o va,or truncado para float
                        n1_truncado_float = float(n1_truncado)

                        #Add o valor a lista  da area da pv1
                        area1.append((n1_truncado_float))


                        #Pega o valor de potência no indice do valor de t  
                        v=pt2[t]

                        #Calculando area da pv1
                        valor_base2=pt2[i]
                        are2=(((v+valor_base2) * tem)/1000)


                        #Trunca os valrod a area2 
                        n2_truncado = Decimal(str(are2)).quantize(Decimal('0.00'), rounding=ROUND_DOWN)

                        #Torna o valor trucado em flaot
                        n2_truncado_float = float(n2_truncado)

                        #Add o valor a lista  da area da pv2
                        area2.append((n2_truncado_float))
                        
                        #Intera sobre t +1
                        t+=1



                except Exception as e:
                        print(e)
                        pass


                print(arquivo)


                #Soma da aréa pv1
                area1_somada=0
                 
                #Somando as areas da pv1
                for valor in  area1:


                    area1_somada+=valor


                #Soma da aréa pv1
                area2_somada=0

                #Somando as areas da pv1
                for valor in  area2:


                    area2_somada+=valor



                #Trunca os valrod a area2 
                n1_truncado = Decimal(str(area1_somada)).quantize(Decimal('0.00'), rounding=ROUND_DOWN)


                #Convertendo o valor truncado apra float
                n1_truncado_float = float(n1_truncado)

                #Pinrtando os valores da area pv1
                print(f'Areá do dia PV1 {n1_truncado_float}')

                print(f'Maior valor da PV1 ={maior_valor_data_1}')

                
                #Add valor da pv1 na lista             
                gerado_pv1.append((n1_truncado_float))

                #Trunca os valrod a area2 
                n2_truncado = Decimal(str(area2_somada)).quantize(Decimal('0.00'), rounding=ROUND_DOWN)

                #Converte valor truncado apra flaot
                n2_truncado_float = float(n2_truncado)

                
                 #Pinrtando os valores da area pv2
                print(f'Areá do dia PV2 {n2_truncado_float}')

                print(f'Maior valor da PV2 ={maior_valor_data_2}')


                #Valor da geração por watts em porcentagem
                val=((n2_truncado_float/580)/(n1_truncado_float/555) -1)*100

                valor_eficiencia=Decimal(str(val)).quantize(Decimal('0.00'), rounding=ROUND_DOWN)

                print(f'Eficiência de geração no dia :{valor_eficiencia}')

                #Add valor da pv2 na lsita               
                gerado_pv2.append((n2_truncado_float))      

                #Add dia e valor de porcentagem a lista
                dias[temporal]=valor_eficiencia


            except:
                 pass
                
                

    except:
    
         pass

#Pega  todas as colunad do DF criado no inicio do código
c = df1.columns

#Cria um indice para interar sobre os exceis
ind=0

#Converte os valors removidos do DIC para uma list
datas = list(dias.keys())
efi = list(dias.values())

#Intera sobre os indices do DIC de dias
for i in range(len(dias)):
     
    #Add valores no DF criado 
    df1 = df1.append({'Dia': datas[ind],'Pv1':gerado_pv1[ind],'Pv2':gerado_pv2[ind],'Eficiência': efi[ind]}, ignore_index=True)
    
    #Intera sobre o indice ind +1
    ind+=1

#Salva o DF criado.
df1.to_excel('Valores de cada mês.xlsx', sheet_name='Meses') 


# In[ ]:




