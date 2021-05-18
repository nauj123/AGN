import imp
from typing import Pattern
from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import os
import spacy
from spacy import matcher
from spacy.matcher import PhraseMatcher
import fitz
from heapq import nlargest
from spacy.lang.es.stop_words import STOP_WORDS
from string import punctuation
import warnings
from pymongo import MongoClient





Client = MongoClient('localhost')

db = Client['tb_sistema_inteligente_agn']
depatarmatetos = db['tb_departamentos']
Municipios = db['tb_municipios']
Actor_Armado = db['tb_actor_armado']
Victima = db['tb_victima']
Enfoque_Diferencial = db['tb_enfoque_diferencial']
Enfoque_Territorial = db['tb_enfoque_territorial']



pdf_file="prueba10.pdf"
pages = convert_from_path(pdf_file,300)
contador = 1 
for page in pages:
    #print(contador)
    filename ="page_"+ str(contador)+".jpg"
    page.save(filename,'JPEG')
    contador+=1


fileimt = contador-1
f = open(pdf_file+".txt","w")

for i in range(1,fileimt+1):
    filename = "page_"+str(i)+".jpg"
    text = str(((pytesseract.image_to_string(Image.open(filename)))))
    text = text.replace('-\n', '')  
    f.write(text)

f.close()


stopwords = list(STOP_WORDS)
nlp = spacy.load('es_core_news_md')
matcher = PhraseMatcher(nlp.vocab)

pattern = nlp("MUNICIPIO")



matcher.add("MUNICIPIO", None,pattern)

doc = nlp(text)
tokens = [token.text.replace('\n','') for token in doc] 
lisd = []
lisM = []
lisA = []
lisV = []
lisE = []
lisT = []

for mathc_id,start,end in matcher(doc):
    span = doc[start:end]
    print("span resultado: ", span.text)


for palabra in doc.ents:
      
      #if palabra.label_ in ['LOC']:

 ### DEPARTAMENTO         
        print(palabra.text,palabra.label_)
        doc_1 = depatarmatetos.find_one({"IN_Codigo_Departamento":palabra.text})
        if doc_1 != None:

             lisd.append(doc_1['IN_Codigo_Departamento'])
             lisd.append(doc_1['IN_Value'])

 ### MUNICIPIO
        doc_1 = Municipios.find_one({"VC_Municipio":palabra.text})
        if doc_1 != None:

             lisM.append(doc_1['VC_Municipio'])
             lisM.append(doc_1['_id']) 

 ### ACTOR ARMADO
        doc_1 = Actor_Armado.find_one({"VC_Tag":palabra.text})     
        if doc_1 != None:

             #lisA.append(doc_1['VC_Tag'])
             lisA.append(doc_1['_id'])
        else: 
            lisA.append(7)      

 ### RESPONSABLE


  ### VICTIMA 
        doc_1 = Victima.find_one({"VC_Tag":palabra.text})       
        if doc_1 != None:

             #lisA.append(doc_1['VC_Tag'])
             lisV.append(doc_1['_id'])
        else: 
            lisV.append(0) 

  ### ENFOQUE DIFERENCIAL               
        doc_1 = Enfoque_Diferencial.find_one({"VC_Tag":palabra.text})       
        if doc_1 != None:

             #lisA.append(doc_1['VC_Tag'])
             lisE.append(doc_1['_id'])
        else: 
            lisE.append(9)

  ### ENFOQUE TERRITORIAL      
        doc_1 = Enfoque_Territorial.find_one({"VC_Tag":palabra.text})       
        if doc_1 != None:

             #lisA.append(doc_1['VC_Tag'])
             lisT.append(doc_1['_id']) 
        else: 
            lisT.append(7)  

     list1 = [10, 20, 10, 30, 40, 40]
print("the unique values from 1st list is")
unique(list1)                        
             

print("DEPARTAMENTO: ",lisd)
print("MUNICIPIOS: ",lisM)
print("ACTOR ARMADO: ",lisA)
print("VICTIMA: ",lisV)
print("ENFOQUE DIFERENCIAL: ",lisE)
print("ENFOQUE TERRITORIAL: ",lisT)


         
         










## verificamos frecuencia .....................................

word_frequencies = {}
for word in doc:
    if word.text.lower() not in stopwords:
        if word.text.lower() not in punctuation:
            if word.text not in word_frequencies.keys():
                word_frequencies[word.text] = 1
            else:
                word_frequencies[word.text] += 1

## maxima frecuancia ···························

max_frequency = max(word_frequencies.values())

# acotamos ......................................

for word in word_frequencies.keys():
    word_frequencies[word] = word_frequencies[word]/max_frequency

# obtiene oraciones ...........................................

sentence_tokens = [sent for sent in doc.sents]

# ordenamos oraciones ........................................ 

sentence_scores = {}
for sent in sentence_tokens:
    for word in sent:
        if word.text.lower() in word_frequencies.keys():
            if sent not in sentence_scores.keys():
                sentence_scores[sent] = word_frequencies[word.text.lower()]
            else:
                sentence_scores[sent] += word_frequencies[word.text.lower()] 

# obtiene 30% del resumen 

select_length = int(len(sentence_tokens)*0.90)
summary = nlargest(select_length, sentence_scores,key=sentence_scores.get)
final_summary = [word.text for word in summary]
summary = ''.join(final_summary)

Salida = open(pdf_file+"_resumen.txt","wb")
Salida.write(summary.encode("utf8"))
Salida.close()


