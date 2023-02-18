import os


os.system('cls' if os.name == 'nt' else 'clear')

lista_archivos = []
print('Buscando archivos .bib en el directorio actual\n')
for arch in os.listdir('./'):
    split_tup = os.path.splitext(arch)
    if split_tup[1] == '.bib':
        print('Archivo encontrado {0}\n'.format(arch))
        lista_archivos.append(arch)


for namearch in lista_archivos:
    have_doi = False
    indexline = 0
    indexlitle = 0
    secuencia = 1
    lines = []
    archivo = open(namearch, 'r', encoding='utf-8')
    for linea in archivo:
        indexline+=1
        linea = linea.strip()
        lines.append(linea+"\n")
        if linea.startswith('title'):
            indexlitle = indexline
        if linea.startswith('doi'):
            have_doi = True
        if linea.startswith('@'):# nuevo articulo
            if not have_doi and indexlitle > 0:
                nueva_linea = "doi={entrysndoi-0000%s},\n" % str(secuencia)
                lines.insert(indexlitle+secuencia,nueva_linea)
                secuencia+=1
            indexlitle = 0
            have_doi = False
    archivo.close()

    with open(namearch, "w", encoding='utf-8') as arch:
        arch.writelines(lines)
