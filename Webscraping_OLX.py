#Webscraping for olx

#0) Set workind directory

#import os
#os.chdir("C:\\Users\\fmgbarbosa\\Documents\\Python\\Projects")

#install : https://datatofish.com/install-package-python-using-pip/

#1) Open browser (chrome)

#import webbrowser

#url = 'www.google.pt'

#chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

#webbrowser.get(chrome_path).open(url)

#2) Download Page info

##import requests
##res = requests.get('https://forecast.weather.gov/MapClick.php?lat=37.78762000000006&lon=-122.39600999999999#.XWQ8MOhKg2w')
##res.raise_for_status() #verificar se existem erros
##
###3) Parse data
##
##import bs4
##objSoup = bs4.BeautifulSoup(res.text,features="html.parser")
##html_obj=objSoup.select('.myforecast-current-lrg')
##temp = html_obj[0].getText()
##
##print(temp)

import requests,bs4,openpyxl

link = 'https://www.olx.pt/telemoveis-e-tablets/telemoveis/'
run = True
row = 2
contador = 0
while run == True:
    res = requests.get(link)
    res.raise_for_status() #verificar se existem erros

    objSoup = bs4.BeautifulSoup(res.text,features="html.parser")

    #get description
    html_obj = objSoup.select('.marginright5 strong')
    comp = len(html_obj)
    comp = comp - 1

    #get price
    price_obj=objSoup.select('.price strong')
    comp = len(price_obj)
    comp = comp - 1

    #get marca for all phones
    link_obj = objSoup.select('.marginright5')
    length = len(link_obj) -1
    lista = []
    for f in range(0,length):
        link_pesq = link_obj[f]['href']
        #get marca
        pagina = requests.get(link_pesq)
        pagina.raise_for_status() #verificar se existem erros
        adSoup = bs4.BeautifulSoup(pagina.text,features="html.parser")
        marca = adSoup.select('td .value')
        label_marca = adSoup.select('th')
        dicionario = {}
        if len(marca) > 0 :
            for t in range(0,len(marca)):
                aux = marca[t].getText()
                value = label_marca[t].getText()
                dicionario[value] = aux.strip()
        lista.append(dicionario) #lista de dicionarios, com a informação de cada anúncio

    #4) Save info into excel file
    #Nota: Para um dado produto o preço e a descrição estão na mesma posição em ambos os vetores.

    #create spreedsheet
    #wb = openpyxl.Workbook()

    #load excel
    wb = openpyxl.load_workbook('example_copy.xlsx')

    #write description and price
    sheet = wb.get_sheet_by_name('Sheet')
    row = sheet.max_row
    j = row
    for i in range (0,comp):
        j = j + 1
        descricao = html_obj[i].getText()
        sheet.cell(row=j,column=1).value = descricao
        sheet.cell(row=j,column=2).value = price_obj[i].getText()
        if 'Anunciante' in lista[i]:
            sheet.cell(row=j,column=3).value = lista[i]['Anunciante']
        else:
            sheet.cell(row=j,column=3).value = 'None'
        if 'Marca' in lista[i]:
            elemen = lista[i]['Marca']
            elemen = elemen.upper()
            sheet.cell(row=j,column=4).value = elemen
        else:
            elemen = 'None'
            sheet.cell(row=j,column=4).value = elemen
        if 'Estado' in lista[i]:
            sheet.cell(row=j,column=5).value = lista[i]['Estado']     
        else:
            sheet.cell(row=j,column=5).value = 'None'
        if 'Sistema Operativo' in lista[i]:
            sheet.cell(row=j,column=6).value = lista[i]['Sistema Operativo']
        else:
            sheet.cell(row=j,column=6).value = 'None'
        if 'Operador' in lista[i]:
            sheet.cell(row=j,column=7).value = lista[i]['Operador']
        else:
            sheet.cell(row=j,column=7).value = 'None'
        #get modelo do iphone
        descricao = descricao.upper()
        lista_descricao = descricao.split(" ")
        if elemen in lista_descricao:
            indice = lista_descricao.index(elemen)
            indice = indice + 1
            if indice <= (len(lista_descricao)-1):
                sheet.cell(row=j,column=8).value = lista_descricao[indice]
        else:
            sheet.cell(row=j,column=8).value = 'None'

    
    wb.save('example_copy.xlsx')

    #5) check new page -> class = pageNextPrev 

    #check for next page
    next_page = objSoup.select('.pageNextPrev')

    if next_page[1].has_attr('href'):
        #se tiver link para a próxima página então repetir ciclo!
        run=True
        #link da página seguintes está no primeiro nivel
        link = next_page[1]['href']
    else:
        run=False
        
    contador = contador + 1
    if contador > 2:
        run = False

