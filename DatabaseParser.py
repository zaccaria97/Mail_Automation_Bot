from email.mime.text import MIMEText 
from email.mime.image import MIMEImage 
from email.mime.application import MIMEApplication 
from email.mime.multipart import MIMEMultipart 
import smtplib 
import time
import imaplib

import os 
import csv
import time
 
# list to store the names of columns
list_of_default_placeholders = []
list_of_column_names = []
listaValoriColonna = []
matriceValoriColonna = [] #Contiene tutti i valori target inseriti dall'utente.
targetRows = [] #Lista contenente le row target
listaOggetto=[] #lista degli oggetti di ciascuna mail da inviare
listaCorpo=[] #lista dei corpi di ciascuna mail da inviare


while(True):
    #lettura del semaforo. il programma si blocca finchè l'utente non clicca su submit dal browser
    with open('semaforo.txt', "r", encoding='Latin1') as semaforo_file:
        flag=''
        while flag!= '1\n':
            flag=semaforo_file.read();
            #print(flag)
            time.sleep(3);
    with open('semaforo.txt', "w", encoding='Latin1') as semaforo_file:
        semaforo_file.flush();
            
       
    with open('OggettoEN.txt') as f:
        oggettoLayoutEN = f.read()
    with open('OggettoDE.txt') as f:
        oggettoLayoutDE = f.read()
    with open('OggettoFR.txt') as f:
        oggettoLayoutFR = f.read()
        
    with open('CorpoEN.txt') as f:
        corpoLayoutEN = f.read()
    with open('CorpoDE.txt') as f:
        corpoLayoutDE = f.read()
    with open('CorpoFR.txt') as f:
        corpoLayoutFR = f.read()



    #Apertura dei placeholder sostitutivi
    with open('DefaultPlaceholders.csv', "r", encoding='Latin1') as csv_file:
        # creating an object of csv reader
        # with the delimiter as ,
        csv_reader = csv.reader(csv_file, delimiter = ',')
     
        # loop to iterate through the rows of csv
        list_of_default_placeholders.clear();
        for row in csv_reader: 
            # adding the first row
            list_of_default_placeholders.append(row)
            # breaking the loop after the
            # first iteration itself
            break
            
    #print(list_of_default_placeholders[0])
    list_of_column_names.clear()
    #Apertura del database
    with open('Database.csv', "r", encoding='Latin1') as csv_file:
        # creating an object of csv reader
        # with the delimiter as ,
        csv_reader = csv.reader(csv_file, delimiter = ',')
     
        # loop to iterate through the rows of csv
        for row in csv_reader: 
            # adding the first row
            list_of_column_names.append(row)
            # breaking the loop after the
            # first iteration itself
            break
            
        # printing the result
        print("Lista delle colonne: ")
        i=0
        for column in list_of_column_names[0]:
            print(column+" | ",end='')
            if column == 'E-mail':
                indiceColonnaEmail=i
            if column == 'Stato/regione':
                indiceColonnaStato=i
            if column == 'Nome':
                indiceColonnaNome=i
            i=i+1
      
      
        #Leggo gli input dell'utente e li inserisco in una matrice.
        matriceValoriColonna.clear()
        targetRows.clear()
        with open('input_utente.txt', "r", encoding='Latin1') as input_utente_file:
            for column in list_of_column_names[0]:
                stringaValoriColonna = input_utente_file.readline();
                stringaValoriColonna2=stringaValoriColonna.strip()
                listaValoriColonna=stringaValoriColonna2.split("|")
                #print("listaValoriColonna= "+str(listaValoriColonna))
                matriceValoriColonna.append(listaValoriColonna)
            for list in matriceValoriColonna:
                print("elemento di matriceValoriColonna= "+str(list))
        
        for row in csv_reader:
            #print(len(row))
            esito=True
            for index in range(len(list_of_column_names[0])):
                #string=str(matriceValoriColonna[index]) +'\n'
                #print(string)
                #print(row[index])
                elementoSenzaNewline=row[index]#+'\n'#.strip()
                
                if not str(matriceValoriColonna[index][0])=='' and not (elementoSenzaNewline in matriceValoriColonna[index]):
                    print("riga non presente");
                    print(elementoSenzaNewline)
                    print(matriceValoriColonna[index])
                    esito=False
                    break
            if esito==True:
                if row[0]!='':
                    targetRows.append(row)
  
  
    header = ['oggetto']+['corpo']+['email'] + list_of_column_names[0]
    index=1
    mettiHeader=True
    print("Numero di row target: "+str(len(targetRows)))
    
    #Se ci sono 0 risultati allora si scrive solo l'intestazione sull'header.
    if len(targetRows)==0: 
        with open('Target.csv', 'w', encoding='UTF8',newline="") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            # write the data
            #writer.writerow(data)
            f.close()
        mettiHeader=False
        
    else:
        for row in targetRows:
            if row[indiceColonnaStato]=='Germania':
                oggettoDaInviare=oggettoLayoutDE
                corpoDaInviare=corpoLayoutDE
            else:
                if row[indiceColonnaStato]=='Francia':
                    oggettoDaInviare=oggettoLayoutFR
                    corpoDaInviare=corpoLayoutFR
                else:
                    oggettoDaInviare=oggettoLayoutEN
                    corpoDaInviare=corpoLayoutEN
                
            index2=0
            for index2 in range(len(list_of_column_names[0])):
                #print(list_of_column_names[0][index2])
                if row[index2]=='':
                    print(index2)
                    oggettoDaInviare=oggettoDaInviare.replace("$$"+str(index2)+"$$",list_of_default_placeholders[0][index2])
                    corpoDaInviare=corpoDaInviare.replace("$$"+str(index2)+"$$",list_of_default_placeholders[0][index2])
                else:
                    if index2==indiceColonnaNome:
                        row[index2] = row[index2].split()[0]
                    oggettoDaInviare=oggettoDaInviare.replace("$$"+str(index2)+"$$",row[index2])
                    corpoDaInviare=corpoDaInviare.replace("$$"+str(index2)+"$$",row[index2])
           
            data=[]
            data.insert(0, oggettoDaInviare)
            print(data[0])
            data.insert(1, corpoDaInviare)
            data.insert(2, row[indiceColonnaEmail])
            data=data+row
            
            #se devo mettere l'header apro in scrittura
            if mettiHeader: 
                with open('Target.csv', 'w', encoding='UTF8',newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    # write the data
                    writer.writerow(data)
                    f.close()
                mettiHeader=False
            else:
                with open('Target.csv', 'a', encoding='UTF8',newline="") as f:
                    writer = csv.writer(f)
                    # write the data
                    writer.writerow(data)
                    f.close()
            