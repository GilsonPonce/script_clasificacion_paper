import pandas as pd
import openpyxl as op
import os

class Article:
    def __init__(self,title:str,abstract:str) -> None:
        self.title = title 
        self.abstract = abstract


class BibFile:
    def __init__(self,file_name:str) -> None:
        self.articles:list[Article] = []
        archivo = open(file_name,'r',encoding='utf-8')
        tmp_article = {'title':'','abstract':''}
        for linea in archivo:
            linea = linea.strip()
            if(linea.startswith('@')): #nuevo articulo
                if(tmp_article['title']=='' or tmp_article['abstract']==''):
                    tmp_article = {'title':'','abstract':''}
                elif (tmp_article['title']!='' and tmp_article['abstract']!=''):
                    self.articles.append(Article(tmp_article['title'],tmp_article['abstract']))
                    tmp_article = {'title':'','abstract':''}
            elif(linea.startswith('title')):
                tmp_title = linea[7:-3]
                tmp_article['title']=tmp_title
            elif(linea.startswith('abstract')):
                tmp_abstract = linea[10:-3]
                tmp_article['abstract']=tmp_abstract
        archivo.close()
        self.articles.append(Article(tmp_article['title'],tmp_article['abstract']))


def generar_matriz_clasificacion(matriz,lista_palabras,list_key_words):
    indice = 0
    for keytitle in lista_palabras:
        indice = 0
        for keyword in list_key_words:
            list_key = keyword.split(',')
            listKey = []
            for palabra in list_key:
                listKey.append(palabra.strip('.').strip())

            for pa in listKey:
                listTwo = pa.split(' ')
                primeraPalabra = listTwo[0]
                tamPa = len(listTwo)
                if tamPa > 1 and primeraPalabra == keytitle:
                    start = lista_palabras.index(keytitle)
                    end = start + tamPa
                    completa = lista_palabras[start:end]
                    if " ".join(completa) == pa:
                        matriz[indice] = 1
                else:
                    if keytitle.strip() == pa:
                        matriz[indice] = 1
            indice += 1

def get_lista_palabras(lista_palabras_articulo):
    lista_palabras = []
    for pal in lista_palabras_articulo:  # escojo cada palabra del titulo
        lista_palabras.append(pal.strip('.').strip(',').strip(';').strip("'").lower())
    return lista_palabras

lista_archivos = []
print('Buscando archivos .bib en el directorio actual\n')
for arch in os.listdir('./'):
    split_tup = os.path.splitext(arch)
    if split_tup[1] == '.bib':
        print('Archivo encontrado {0}'.format(arch))
        lista_archivos.append(arch)

#Lee el archivo de categorias
print('Cargando Categorias y Key words del archivo Field categories.xlsx')
excel_data_df = pd.read_excel("Field categories.xlsx")
list_category = excel_data_df['Indicator name'].tolist() #extraigo los indicadores
list_key_words = excel_data_df['Key words'].tolist() # extraigo una lista de los keywords de cada indicador

for archivo in lista_archivos:
    list_paper = [] #contienen los titulos de cada paper
    matriz_general = [] #matriz de clasificacion padre
    bib = BibFile(archivo)
    for articulo in bib.articles:
        print('Leyendo articulo {0}'.format(articulo.title))
        list_paper.append(articulo.title) #agrego a la lista el titulo
        #Obtiene cada palbara del titulo
        list_key_title = get_lista_palabras(articulo.title.split(' '))
        #Obtiene cada palabra del abstract
        list_key_abstract = get_lista_palabras(articulo.abstract.split(' '))
        #Genera matriz de clasificacion
        print('Generando la matriz de clasificacion............')
        matriz = [0 for i in range(len(list_category))]#matriz de ceros para ese articulo
        generar_matriz_clasificacion(matriz,list_key_title,list_key_words)
        generar_matriz_clasificacion(matriz,list_key_abstract,list_key_words)
        matriz_general.append(matriz)

    # Arma el archivo de excel con la matriz padre de clasificacion
    wb = op.Workbook()
    hoja = wb.active
    hoja.title = "Clasificacion"
    hoja.append(('Titulos Papers',)+tuple(list_category))
    for i in range(len(list_paper)):
        hoja.append((list_paper[i],)+tuple(matriz_general[i]))
    print('Generando Archivo de clasificacion {0}'.format(ar[0] + '.xlsx'))
    ar = archivo.split('.')
    wb.save(ar[0]+'.xlsx')

print('############ Archivos de clasificacion Generados Exitosamente! ############')