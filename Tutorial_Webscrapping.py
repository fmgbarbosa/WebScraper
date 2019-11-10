#install packages

#get packages
import requests,bs4,openpyxl,os

#Download da página
res = requests.get('https://www.custojusto.pt/porto/telefones-acessorios')
print(type(res))
print(res.text)

res.raise_for_status()

#save page

#file = open('pag_html.txt','wb')
#for text in res.iter_content(100000):
#    file.write(text)
#file.close()

#openpage

res = open('pag_html.txt')

#Parse da página HTML
objSoup = bs4.BeautifulSoup(res,features="html.parser")
print(type(objSoup))
#objSoup=bs4.BeautifulSoup(res.text,features="html.parser")

#Get values
#.title_related b
vetor_ad_name = objSoup.select('.title_related b')
vetor_ad_price = objSoup.select('.container_related .price_related')
print(len(vetor_ad_name))
print(len(vetor_ad_price))

#Save values into dictionary
list_ad = [] #vetor com anuncios
comp = len(vetor_ad_name)-1
for i in range (0,comp):
    ad = {} #anuncio
    ad["title"] = vetor_ad_name[i].getText(strip=True)
    ad["price"] = vetor_ad_price[i].getText(strip=True)
    list_ad.append(ad)

#Save in excel

#set directory
os.chdir('C:\\Users\\Francisco Barbosa\\Documents\\Scripts_Python')

#create excel file
workbook = openpyxl.Workbook()
sheet = workbook.get_sheet_by_name('Sheet')
sheet.title = 'anuncios'
row = 0
for f in list_ad:
    row = row + 1
    sheet.cell(row = row,column = 1).value = f['title']
    sheet.cell(row=row, column = 2).value = f['price']

workbook.save('tutorial.xlsx')

