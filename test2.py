import pandas as pd
import openpyxl as op
import os
import errno
import sys

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

def buscar_paper_sin_doi(namearch,secuencia=1):
    list_doi = []
    if namearch == '':
        return list_doi
    info_para_excel = []
    have_doi = False
    indexline = 0
    indexlitle = 0
    conteo_por_archivo = 0
    lines = []
    totalLineas = 0
    conteo_lineas =0
    with open(namearch, "r", encoding='utf-8') as arch:
        for linea in arch:
            totalLineas += 1
    archivo = open(namearch, 'r', encoding='utf-8')
    for linea in archivo:
        conteo_lineas += 1
        indexline += 1
        linea = linea.strip()
        lines.append(linea + "\n")
        if linea.startswith('title'):
            indexlitle = indexline
            info_para_excel.append(linea[7:-2])
            info_para_excel.append(namearch)
        if linea.startswith('doi'):
            have_doi = True
            info_para_excel.append(linea[5:-2])
            info_para_excel.reverse()
            list_doi.append(tuple(info_para_excel))
        if linea.startswith('@') or conteo_lineas == totalLineas:  # nuevo articulo
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
    return list_doi,secuencia

def generar_excel(encabezado=[],contenido={},name_archivo='archivo'):
    if len(encabezado) == 0:
        return
    if len(contenido) == 0:
        return
    wb = op.Workbook()
    for key in contenido:
        if key == 'papers':
            hoja = wb.active
            hoja.title = "CLASIFICACION"
            hoja.append(encabezado[0])
            for i in range(len(contenido[key])):
                hoja.append(contenido[key][i])
        if key == 'dois':
            hojadois = wb.create_sheet("DOIS")
            hojadois.append(encabezado[1])
            hojadois.column_dimensions['A'].width = 36
            hojadois.column_dimensions['B'].width = 19
            hojadois.column_dimensions['C'].width = 249
            for i in range(len(contenido[key])):
                hojadois.append(contenido[key][i])
    print('### Generando Archivo {0}\n'.format(name_archivo + '.xlsx'))
    directorio_matrices = 'MATRICES'
    crear_directorio(directorio_matrices)
    wb.save('./' + directorio_matrices + '/' + name_archivo + '.xlsx')

def escribir_archivo_ultimo_doi_generado(secuencia):
    archivo = open("ultimodoi.txt","w")
    archivo.write(str(secuencia))
    archivo.close()

def leer_archivo_ultimo_doi_generado():
    ruta_archivo = "./ultimodoi.txt"
    secuencia = 0
    if os.path.exists(ruta_archivo):
        archivo = open("ultimodoi.txt","r")
        secuencia_arch = archivo.readline().strip()
        if secuencia_arch :
            secuencia = int(secuencia_arch)
    return secuencia

os.system('cls' if os.name == 'nt' else 'clear')
lista_archivos = []
print('Buscando archivos .bib en el directorio actual\n')
for arch in os.listdir('./'):
    split_tup = os.path.splitext(arch)
    if split_tup[1] == '.bib':
        print('Archivo encontrado {0}\n'.format(arch))
        lista_archivos.append(arch)

#Lee el archivo de categorias
print('Cargando Categorias y Key words del archivo Field categories.xlsx\n')
excel_data_df = pd.read_excel("Field categories.xlsx")
list_category = excel_data_df['Indicator name'].tolist() #extraigo los indicadores
list_key_words = excel_data_df['Key words'].tolist() # extraigo una lista de los keywords de cada indicador

secuencia = 1

if len(sys.argv) > 1:
    numero = sys.argv[1]
    ultimo = int(numero)
    secuencia = ultimo + 1

secuencia_archivo = leer_archivo_ultimo_doi_generado()
if secuencia_archivo != 0 and len(sys.argv) <= 1:
    secuencia = secuencia_archivo + 1

for archivo in lista_archivos:
    list_paper = [] #contienen los titulos de cada paper
    matriz_general = [] #matriz de clasificacion padre
    print('########## BUSCANDO PAPER SIN DOI ###########')
    lista_doi, ultima_secuencia = buscar_paper_sin_doi(archivo,secuencia)
    if ultima_secuencia == secuencia:
        escribir_archivo_ultimo_doi_generado(secuencia= ultima_secuencia-1)
    else:
        escribir_archivo_ultimo_doi_generado(secuencia=ultima_secuencia)
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
    encabezado_excel = [('Titulos Papers',)+tuple(list_category)]
    list_contenido_excel = []
    ar = archivo.split('.')
    for i in range(len(list_paper)):
        list_contenido_excel.append((list_paper[i],)+tuple(matriz_general[i]))
    diccionario_contenido = {}
    diccionario_contenido['papers'] = list_contenido_excel
    if (lista_doi) != 0:
        encabezado_excel.append(('Doi','Nombre del Archivo','Titulo del Paper'))
        diccionario_contenido['dois'] = lista_doi
    generar_excel(encabezado=encabezado_excel,contenido=diccionario_contenido,name_archivo=ar[0])

print('############ Archivos de clasificacion Generados Exitosamente! ############\n')

      
        