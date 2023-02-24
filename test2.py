import pandas as pd
import openpyxl as op
import os
import errno

class Article:
    def __init__(self,title:str,abstract:str) -> None:
        self.title = title 
        self.abstract = abstract

class BibFile:
    def __init__(self,file_name:str) -> None:
        self.articles:list[Article] = []
        archivo = open(file_name,'r',encoding='utf-8')
        tmp_article = {'title':'','abstract':''}
        have_doi = False
        for linea in archivo:
            linea = linea.strip()
            if(linea.startswith('@')): #nuevo articulo
                if(tmp_article['title']=='' or tmp_article['abstract']==''):
                    tmp_article = {'title':'','abstract':''}
                elif (tmp_article['title']!='' and tmp_article['abstract']!=''):
                    self.articles.append(Article(tmp_article['title'],tmp_article['abstract']))
                    tmp_article = {'title':'','abstract':''}
            elif(linea.startswith('title')):
                tmp_title = linea[7:-2]
                tmp_article['title']=tmp_title
            elif(linea.startswith('abstract')):
                tmp_abstract = linea[10:-2]
                tmp_article['abstract']=tmp_abstract
            elif(linea.startswith('doi')):
                have_doi = True

        archivo.close()
        self.articles.append(Article(tmp_article['title'],tmp_article['abstract']))


def generar_matriz_clasificacion(matriz,lista_palabras,list_key_words):
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

def crear_directorio(nombre):
    try:
        os.mkdir(nombre)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise

def buscar_paper_sin_doi(list_archivos_bib):
    list_doi = []
    if len(list_archivos_bib) == 0:
        return list_doi
    secuencia = 1
    for namearch in list_archivos_bib:
        info_para_excel = []
        have_doi = False
        indexline = 0
        indexlitle = 0
        conteo_por_archivo = 0
        lines = []
        archivo = open(namearch, 'r', encoding='utf-8')
        for linea in archivo:
            indexline += 1
            linea = linea.strip()
            lines.append(linea + "\n")
            if linea.startswith('title'):
                indexlitle = indexline
                info_para_excel.append(linea[7:-2])
                info_para_excel.append(namearch)
            if linea.startswith('doi'):
                have_doi = True
            if linea.startswith('@'):  # nuevo articulo
                if not have_doi and indexlitle > 0:
                    newdoi = "entrysndoi-0000%s" % str(secuencia)
                    nueva_linea = "doi={%s},\n" % newdoi
                    lines.insert(indexlitle + conteo_por_archivo, nueva_linea)
                    secuencia += 1
                    conteo_por_archivo += 1
                    info_para_excel.append(newdoi)
                    info_para_excel.reverse()
                    list_doi.append(tuple(info_para_excel))
                indexlitle = 0
                have_doi = False
                info_para_excel.clear()
        archivo.close()
        with open(namearch, "w", encoding='utf-8') as arch:
            arch.writelines(lines)
    return list_doi

def generar_excel(encabezado=(),contenido=[],titulo='Clasificacion',directorio='matrices',name_archivo='newsDoi'):
    wb = op.Workbook()
    hoja = wb.active
    hoja.title = titulo
    hoja.append(encabezado)
    for i in range(len(contenido)):
        hoja.append(contenido[i])
    print('### Generando Archivo {0}\n'.format(name_archivo + '.xlsx'))
    directorio_matrices = directorio
    crear_directorio(directorio_matrices)
    wb.save('./' + directorio_matrices + '/' + name_archivo + '.xlsx')


os.system('cls' if os.name == 'nt' else 'clear')
lista_archivos = []
print('Buscando archivos .bib en el directorio actual\n')
for arch in os.listdir('./'):
    split_tup = os.path.splitext(arch)
    if split_tup[1] == '.bib':
        print('Archivo encontrado {0}\n'.format(arch))
        lista_archivos.append(arch)

print('########## BUSCANDO PAPER SIN DOI ###########')
lista_doi_nuevos = buscar_paper_sin_doi(lista_archivos)
if len(lista_doi_nuevos) > 0:
    encabezado_excel_doi = ('Nuevo Doi','Nombre del Archivo','Titulo del Paper')
    generar_excel(encabezado=encabezado_excel_doi,contenido=lista_doi_nuevos,titulo="NuevosDOI",directorio='DOI')

#Lee el archivo de categorias
print('Cargando Categorias y Key words del archivo Field categories.xlsx\n')
excel_data_df = pd.read_excel("Field categories.xlsx")
list_category = excel_data_df['Indicator name'].tolist() #extraigo los indicadores
list_key_words = excel_data_df['Key words'].tolist() # extraigo una lista de los keywords de cada indicador

for archivo in lista_archivos:
    list_paper = [] #contienen los titulos de cada paper
    matriz_general = [] #matriz de clasificacion padre
    bib = BibFile(archivo)
    for articulo in bib.articles:
        print('Leyendo articulo {0}\n'.format(articulo.title))
        list_paper.append(articulo.title) #agrego a la lista el titulo
        #Obtiene cada palbara del titulo
        list_key_title = get_lista_palabras(articulo.title.split(' '))
        #Obtiene cada palabra del abstract
        list_key_abstract = get_lista_palabras(articulo.abstract.split(' '))
        #Genera matriz de clasificacion
        print('Generando la matriz de clasificacion............\n')
        matriz = [0 for i in range(len(list_category))]#matriz de ceros para ese articulo
        generar_matriz_clasificacion(matriz,list_key_title,list_key_words)
        generar_matriz_clasificacion(matriz,list_key_abstract,list_key_words)
        matriz_general.append(matriz)

    # Arma el archivo de excel con la matriz padre de clasificacion
    encabezado_excel = ('Titulos Papers',)+tuple(list_category)
    list_contenido_excel = []
    ar = archivo.split('.')
    for i in range(len(list_paper)):
        list_contenido_excel.append((list_paper[i],)+tuple(matriz_general[i]))
    generar_excel(encabezado=encabezado_excel,contenido=list_contenido_excel,name_archivo=ar[0])

print('############ Archivos de clasificacion Generados Exitosamente! ############\n')

      
        