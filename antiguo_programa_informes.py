from pathlib import Path
from turtle import position
import win32com.client
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import tkinter as tk
import win32serviceutil
import win32service
import win32event



############################################# INFORMACION GENERAL #########################################

#Variables globales
TIPO="PRO"

#Diccionario con las categorias de las aplicaciones 
categories = dict(Websphere =".w61",BBDD =".BD",NET=".NET",ClientServidor=".NT",BIM=".BIM",Documentum=".doc",Portlet=".plr",Devops=["DevOps","execució scripts BD","desplegament"],
Cognos=[".CBI",".CDM"],Paquet=["Distribució","Assignació"])

#Dataset donde se guardan los datos
dataset = pd.DataFrame(columns=['Aplicacion','Tecnologia','Resultado','Urgente','Error','Mes','Fecha'])
dataset_totales = pd.DataFrame(columns=['Aplicacion','Tecnologia','Resultado','Urgente','Error','Mes','Fecha'])




#Lista con nombre de los meses que contiene el numero de semanas.
MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
#Lista con el numero de semanas que contiene cada mes
SEMANAS = [4,4,5,4,4,5,4,4,4,4,4,4]

#Generar lista Enero=["Semana 1","Semana 2","Semana 3","Semana 4"]
lista_meses = []
for i in range(len(MESES)):
    lista_meses.append(["Semana " + str(x) for x in range(1,SEMANAS[i]+1)])



######################################## CONECCION AL CORREO ###############################################

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

############################################################################################################
    
def generate_data(messages,type,mouth,week):
    #create the output folder
    output_folder = Path.cwd() /'Informes Anuales'/type/mouth/week
    output_folder.mkdir(parents=True,exist_ok=True)

    #Loop through all messages
    for message in messages:

        #get the subject
        subject=message.Subject

        #get the categories
        categories_list = message.Categories

        application =subject
        tecnologia = "No definida"
        
        #get the Error, is the last text in the subject before the Error word
        #if the subject doesn't contain the word Error, the error is "No definido"
        
        error = ""
        
        for key in categories:
            #if the value is a list then we will iterate through the list
            if isinstance(categories[key],list):
                for value in categories[key]:
                    if value in subject:
                        tecnologia = key
                        break
                 
            elif categories[key] in subject:
                    #Descomnentar si solo se quiere ver el nombre de la aplicacion.
                    #application = subject.split(categories[key])[0].split()[-1]
                    tecnologia = key
                    break
            


        #If the subject start with the word "URGENT" the application is urgent
        if subject.startswith("URGENT"):
            urgente = "SI"
        else:
            urgente = "NO"

        #if the list of categories contains the word "KO" the result is KO
        if "KO" in categories_list:
            resultado = "KO"
            error = subject.split("Error:")[-1]
        else:
            resultado = "OK"
            


        #is the message contain the word "Scripts per homologar" add an entry to the dataset
        if "Scripts per homologar" in message.Body and tecnologia != "BBDD" and "Tornar enrere" not in message.Subject:
            tecnologia2 = "BBDD"
            dataset.loc[len(dataset)] = [application,tecnologia2,resultado,urgente,error,mouth,message.SentOn.strftime("%d/%m/%Y")]
            dataset_totales.loc[len(dataset_totales)] = [application,tecnologia2,resultado,urgente,error,mouth,message.SentOn.strftime("%d/%m/%Y")]



        dataset.loc[len(dataset)] = [application,tecnologia,resultado,urgente,error,mouth,message.SentOn.strftime("%d/%m/%Y")]
        dataset_totales.loc[len(dataset_totales)] = [application,tecnologia,resultado,urgente,error,mouth,message.SentOn.strftime("%d/%m/%Y")]
        #set all the variable to None
        print(dataset)


    #el separador es la tabulacion
    dataset.to_csv(output_folder / f'{mouth}_{week}.csv',index=False,encoding='utf-8-sig')

    
    dataset.groupby(['Tecnologia','Resultado']).size().unstack().plot(kind='bar',stacked=True)

    #save the plot to a png file
    plt.savefig(output_folder / f'{mouth}_{week}.png',bbox_inches='tight')

    #clean the dataset
    dataset.drop(dataset.index,inplace=True)
    
    dataset_totales.to_csv('InformesTotales'+type+'.csv',index=False,encoding='utf-8-sig')
   
    return dataset_totales


####################################################################################################
'''continue
Funcion que rocogue la informacion de los correos y la guarde en un dataset.
Utiliza la variable global TIPO para saber si es un informe de produccion o preproduccion.
'''
####################################################################################################

def get_data():
    for mes in MESES:
        for semana in lista_meses[MESES.index(mes)]:
            #Tomamos la informacion de la carpetas carpetas siguiendo esta estructura
            #try to get the folder
            try:
                inbox=outlook.Folders("gestioversions@bcn.cat").Folders("Bandeja de entrada").Folders("2022").Folders(mes).Folders(semana).Folders(TIPO)
            except:
                print("No existe la carpeta " + semana + " en el " + mes)
                continue
          
            #get all the messages from the inbox
            messages = inbox.Items
            if inbox.Items.Count > 0:
                generate_data(messages,TIPO,mes,semana)

def get_data_month(mes):
    for semana in lista_meses[MESES.index(mes)]:
        #Tomamos la informacion de la carpetas carpetas siguiendo esta estructura
        #try to get the folder
        try:
            inbox=outlook.Folders("gestioversions@bcn.cat").Folders("Bandeja de entrada").Folders("2022").Folders(mes).Folders(semana).Folders(TIPO)
        except:
            print("No existe la carpeta " + semana + " en el " + mes)
            continue

        #get all the messages from the inbox
        messages = inbox.Items
        if inbox.Items.Count > 0:
            return generate_data(messages,TIPO,mes,semana)
    

def get_data_week(mes,semana):
    inbox=outlook.Folders("gestioversions@bcn.cat").Folders("Bandeja de entrada").Folders("2022").Folders(mes).Folders(semana).Folders(TIPO)
    messages = inbox.Items
    if inbox.Items.Count > 0:
         return generate_data(messages,TIPO,mes,semana)

 
####################################################################################################
'''
Funcion que genera los me de los informes, agrupa por tecnologia y resultado y los guarda en un png.
Esta informacion es relativa al informe total por año.
'''
####################################################################################################

def group_by_tecnologies(dataset_totales):
    #quit the none values from the dataset_totales
    dataset_totales = dataset_totales[dataset_totales.Tecnologia != "No definida"]
    
    #set the labels
    labels = dataset_totales.pivot_table(index=['Tecnologia'], aggfunc='size').index

    #quit empty values from labels
    labels = [x for x in labels if x != ""]

    #set the colors, get a professional color palette 
    colors = ['#f94144','#f3722c','#f8961e','#f9844a','#f9c74f','#90be6d','#43aa8b','#4d908e','#577590','#277da1']

    #set the figure size
    plt.figure(figsize=(12,12))

    dataset_totales.groupby(['Tecnologia']).size().plot(kind='pie',autopct='%1.1f%%',colors=colors)

    #set the tittle
    plt.title("Porcentaje de incidencies per tecnologia en "+TIPO)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("")

    plt.legend(labels,loc="upper right")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_tecnologias.png',bbox_inches='tight')


####################################################################################################
'''
Funcion que genera los plots de los informes, agrupa por tecnologia y resultado y los guarda en un png.
Esta informacion es referente al mes entero.
'''
####################################################################################################

def group_by_tecnologies_month(dataset_totales, month):
    #quit the none values from the dataset_totales
    dataset_totales = dataset_totales[dataset_totales.Tecnologia != "No definida"]
    
    #set the labels
    labels = dataset_totales.pivot_table(index=['Tecnologia'], aggfunc='size').index

    #quit empty values from labels
    labels = [x for x in labels if x != ""]

    #set the colors, get a professional color palette 
    colors = ['#f94144','#f3722c','#f8961e','#f9844a','#f9c74f','#90be6d','#43aa8b','#4d908e','#577590','#277da1']

    #set the figure size
    plt.figure(figsize=(12,12))

    dataset_totales[dataset_totales.Mes == month].groupby(['Tecnologia']).size().plot(kind='pie',autopct='%1.1f%%',colors=colors)

    #set the tittle
    plt.title("Porcentaje de incidencies per tecnologia en "+TIPO+" en el mes de "+month)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("")

    plt.legend(labels,loc="best")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_tecnologias_'+month+'.png',bbox_inches='tight')


####################################################################################################
'''

'''
####################################################################################################

def result_by_month_ok_ko(dataset_totales):
    #do not order by the alphabet but by the month
    dataset_totales['Mes'] = pd.Categorical(dataset_totales['Mes'], categories=MESES, ordered=True)

    #get a graph of the evolution of the month of the incidents with OK and KO set colors
    dataset_totales.groupby(['Mes','Resultado']).size().unstack().plot(kind='bar',color=['#f94144','#90be6d'])

    #set the number of OK and KO by month in the plot
    for p in plt.gca().patches:
        if p.get_height() > 0:
            plt.gca().annotate('{:.0f}'.format(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()), ha = 'center', va = 'center', 
            xytext = (0, 10), textcoords = 'offset points')

    #set the tittle
    plt.title("Evolució de les incidències per mes en "+TIPO)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("Número d'incidències")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_evolution.png',bbox_inches='tight')


####################################################################################################
'''

'''
####################################################################################################

def relation_tecnologies_resultado(dataset_totales,mes):
    #get the graph of numbers of incidents by tecnologies and result
    dataset_totales[dataset_totales.Mes == mes].pivot_table(index=['Tecnologia'],columns=['Resultado'],aggfunc='size').plot(kind='bar',color=['#f94144','#90be6d'])

    #set the number of OK and KO by tecnology in the plot
    for p in plt.gca().patches:
        if p.get_height() > 0:
            plt.gca().annotate('{:.0f}'.format(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()), ha = 'center', va = 'center', 
            xytext = (0, 10), textcoords = 'offset points')

    #set the tittle
    plt.title("Incidències per tecnologia i resultat en "+TIPO+" en el mes de "+mes)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("Numero d'incidencies")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_tecnologies_resultado_'+mes+'.png',bbox_inches='tight')
    
####################################################################################################
'''

'''
####################################################################################################

def relation_incidencies_urgent(dataset_totales):
    dataset_totales['Mes'] = pd.Categorical(dataset_totales['Mes'], categories=MESES, ordered=True)

    #get the numbers of incidents  if is an urgent or not by month
    dataset_totales.pivot_table(index=['Mes'],columns=['Urgente'],aggfunc='size').plot(kind='bar',color=['#90be6d','#f94144'])

    #set the number of urgent and not urgent by month in the plot
    for p in plt.gca().patches:
        if p.get_height() > 0:
            plt.gca().annotate('{:.0f}'.format(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()), ha = 'center', va = 'center', 
            xytext = (0, 10), textcoords = 'offset points')

    #set the tittle
    plt.title("Incidències per mes i urgència en "+TIPO)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("Numero d'incidencies")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_incidencies_urgent.png',bbox_inches='tight')
    


def relation_incidencies_urgent_month(dataset_totales, month):
    #get the numbers of incidents  if is an urgent or not by month
    dataset_totales[dataset_totales.Mes == month].pivot_table(index=['Tecnologia'],columns=['Urgente'],aggfunc='size').plot(kind='bar',color=['#90be6d','#f94144'])

    #set the number of urgent and not urgent by month in the plot
    for p in plt.gca().patches:
        if p.get_height() > 0:
            plt.gca().annotate('{:.0f}'.format(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()), ha = 'center', va = 'center', 
            xytext = (0, 10), textcoords = 'offset points')

    #set the tittle
    plt.title("Incidències per tecnologia i urgència en "+TIPO+" en el mes de "+month)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("Numero d'incidencies")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_incidencies_urgent_'+month+'.png',bbox_inches='tight')



####################################################################################################
'''

'''
####################################################################################################

def incidences_by_month_PRE_and_PRO(dataset_pre,dataset_pro):
    
    #add a columns=['tipo'] to the dataset of the PRE and PRO
    dataset_pre['Tipo'] = 'PRE'
    dataset_pro['Tipo'] = 'PRO'

    #concat the two datasets
    dataset_totales = pd.concat([dataset_pre,dataset_pro])
    dataset_totales['Mes'] = pd.Categorical(dataset_totales['Mes'], categories=MESES, ordered=True)

    #graph the number of incidents by month for PRE and PRO
    dataset_totales.pivot_table(index=['Mes'],columns=['Tipo'],aggfunc='size').plot(kind='bar',color=['#f8961e','#4d908e'])

    #set the number of PRE and PRO by month in the plot
    for p in plt.gca().patches:
        if p.get_height() > 0:
            plt.gca().annotate('{:.0f}'.format(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()), ha = 'center', va = 'center', 
            xytext = (0, 10), textcoords = 'offset points')

    #set the tittle
    plt.title("Incidències per mes i tipus en "+TIPO)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("Numero d'incidencies")

    #save the plot to a png file
    plt.savefig('Report_'+TIPO+'_incidencies_PRE_PRO.png',bbox_inches='tight')
   
  



####################################################################################################
'''

'''
####################################################################################################

def get_plots(TIPO,dataset):
    #create the folder plots if not exist or change to the folder if exist
    if not os.path.exists('plots'):
        os.makedirs('plots')
        #if not exist the folder TIPO create it
    if not os.path.exists('plots/'+TIPO):
        os.makedirs('plots/'+TIPO)


    os.chdir('plots/'+TIPO)

    group_by_tecnologies(dataset)
    relation_incidencies_urgent(dataset)
    result_by_month_ok_ko(dataset)

    
    for mes in MESES:
        #create a folder to save the plots if not exist with the name of the month
        if not os.path.exists(mes):
            os.makedirs(mes)
        #change the directory to the folder plots

        #try and only generate the data if the month exist in the dataset
        try:
            os.chdir(mes)
            group_by_tecnologies_month(dataset,mes)
            relation_incidencies_urgent_month(dataset,mes)
            relation_tecnologies_resultado(dataset,mes)
            os.chdir('..')
        except:
            print("No existe el mes "+mes+" en el dataset")

def get_plot_mes_semana(TIPO,dataset,mes,semana):
    #create the folder plots if not exist or change to the folder if exist
    if not os.path.exists('plots'):
        os.makedirs('plots')
        #if not exist the folder TIPO create it
    if not os.path.exists('plots/'+TIPO):
        os.makedirs('plots/'+TIPO)


    os.chdir('plots/'+TIPO)

    #create a folder to save the plots if not exist with the name of the month
    if not os.path.exists(mes):
        os.makedirs(mes)
    #change the directory to the folder plots

    #try and only generate the data if the month exist in the dataset
    try:
        os.chdir(mes)
        #create a folder to save the plots if not exist with the name of the week
        if not os.path.exists(semana):
            os.makedirs(semana)
        os.chdir(semana)

        group_by_tecnologies_month(dataset,mes)
        relation_incidencies_urgent_month(dataset,mes)
        relation_tecnologies_resultado(dataset,mes)
       
    except:
        print("No existe el mes "+mes+" en el dataset")

####################################################################################################
'''

'''
####################################################################################################

def generate_doc_resum(archivo):
    #Create a pdf file with the resum of the incidents
    #put the folling titteles in the pdf
    titulos = ['Mes','Tecnologia','Resultado','Urgente','Descripcion','Solucion','Observaciones']
    #create a dataframe with the titteles
    df = pd.DataFrame(columns=titulos)
    #make a table with the following columns, "Tecnologia","Produccion Ok","Produccion KO","Total Produccion"
    df2 = pd.DataFrame(columns=['Tecnologia','Produccion Ok','Produccion KO','Total Produccion'])

    #get the number of incidents by tecnology
    df2['Tecnologia'] = archivo['Tecnologia'].value_counts().index
    df2['Total Produccion'] = archivo['Tecnologia'].value_counts().values

    #get the number of incidents by tecnology and result
    df3 = archivo.pivot_table(index=['Tecnologia'],columns=['Resultado'],aggfunc='size')
    #change the NaN values to 0
    df3 = df3.fillna(0)

    #get the number of incidents by tecnology and result

    df2['Produccion Ok'] = df3['OK'].values
    df2['Produccion KO'] = df3['KO'].values

    #add a last row with the total of the incidents
    df2.loc[len(df2)] = ['Total',df2['Produccion Ok'].sum(),df2['Produccion KO'].sum(),df2['Total Produccion'].sum()]

 
    #generate a plot with the dataset df2
    df2.plot(x='Tecnologia',y=['Produccion Ok','Produccion KO'],kind='bar',color=['#90be6d','#f94144'])


    #set the tittle
    plt.title("Incidències per tecnologia i resultat en "+TIPO)

    #set the left title this is to aboid the problem with the tittle
    plt.ylabel("Numero d'incidencies")

    #save the plot to a png file
    plt.savefig('Final_Report.png',bbox_inches='tight')


    #print the information
    print(df2)
    


def main(mes="mes"):
    #solicitar al usuario si desea generar el reporte de todos los meses o de un mes en especifico
    #si el usuario desea generar el reporte de todos los meses
    if mes == 'todos':
        
        get_data()

        #read csv file
        dataset= pd.read_csv('InformesTotales'+TIPO+'.csv',sep=',',encoding='utf-8-sig')

        get_plots(TIPO,dataset)

    
    #si el usuario desea generar el reporte de un mes en especifico y una semana en especifico
    elif mes == 'mes':
       
        print("Seleccione el mes que desea generar el reporte")
        for i in range(len(MESES)):
            print(str(i+1)+" - "+MESES[i])

        try:
            while True:
                mes = input('Introduce el mes: ')
                mes = int(mes)
                if MESES[mes-1] in MESES:
                    mes= MESES[mes-1]
                    break
                else:
                    print('El mes no es valido, por favor introduce un mes valido')
            
            
            for element in lista_meses[MESES.index(mes)]:
                print(str(lista_meses[MESES.index(mes)].index(element)+1)+" - "+element)

            while True:
                semana = input('Introduce la semana: ')
                semana = int(semana) - 1
                if lista_meses[MESES.index(mes)][semana] in lista_meses[MESES.index(mes)]:
                    semana = lista_meses[MESES.index(mes)][semana]
                    break
                else:
                    print('La semana no es valida, por favor introduce una semana valida')
        
        except:
            print("El mes o la semana no es valido, por favor introduce un valor valido")
            main()
                

        #get the data
        dataset_mes_semana = get_data_week(mes,semana)
        get_plot_mes_semana(TIPO,dataset_mes_semana,mes,semana)
        #save the data to a csv file
        dataset_mes_semana.to_csv('Informes'+TIPO+'_'+mes+'_'+semana+'.csv',sep=',',encoding='utf-8-sig',index=False)


if __name__ == "__main__":
    

    print("=====================================================================================================")
    print("Generador de informes de peticiones")
    print("=====================================================================================================")

    print("Bienvenido al generador de informes de peticiones")
    print("Este programa genera un informe de las peticiones realizadas en el mes seleccionado o de todos los meses")
    print("=====================================================================================================")

    print("Primero deberas instalar las librerias necesarias")
    print("=====================================================================================================")


    if input("¿Desea instalar las librerias? (y/n): ") == 'y':
        #instalar los requrimientos del archivo requirements.txt
        os.system('pip install -r requirements.txt')

    print("¿Desea generar el reporte de todos los meses o de un mes en especifico?")
    print("1. Todos los meses")
    print("2. Un mes en especifico")
    print("3. Salir")

    opcion = input("Introduzca la opcion: ")
    if opcion == '1':
        main('todos')
    elif opcion == '2':
        main('mes')
    elif opcion == '3':
        exit()




    




