import pandas as pd
import openpyxl
import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#from IPython.display import HTML
import win32com.client as client

vendas = pd.read_excel('vendas.xlsx')

def formatar(valor):
     return "R${:,.2f}".format(valor)

faturamento_top_25 = vendas.groupby('ID Loja')[['Valor Final','Quantidade']].sum().sort_values(by='Valor Final',ascending=False).head(25)
faturamento_top_25['Ticket Médio'] = faturamento_top_25['Valor Final']/faturamento_top_25['Quantidade']
faturamento_top_25['Ticket Médio'] = round(faturamento_top_25['Ticket Médio'],2).apply(formatar)
faturamento_top_25['Valor Final'] = faturamento_top_25['Valor Final'].apply(formatar)
faturamento_top_25['Email']= 'paloma.moozer@gmail.com'
ranking = faturamento_top_25[['Valor Final', 'Quantidade','Ticket Médio']].to_html()


for indice, row in faturamento_top_25.iterrows():
    destinatario = row['Email']
    faturamento = row['Valor Final']
    qntd_vendida = row['Quantidade']
    ticket_medio = row['Ticket Médio']
    loja = indice 

    corpo_email = f"""
     <p>Segue relatório da sua loja {loja}</p>
     <p>Faturamento = {faturamento}</p>
     <p>Quantidade vendida = {qntd_vendida}</p>
     <p>Ticket Médio = {ticket_medio}</p>
     """
    msg = email.message.Message()
    msg['Subject']= 'Relatório semanal'
    msg['From']= 'paloma.moozer@gmail.com'
    msg['To']= destinatario
    password = 'chrrietkbuchjxhm'
    msg.add_header('Content-Type','text/html')

    msg.set_payload(corpo_email)

    s= smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()

    s.login(msg['From'],password)
    s.sendmail(msg['From'],[msg['To']], msg.as_string().encode('utf-8'))

#print(ticket_medio_ig_campinas)

def enviar_email():
     corpo_email = f"""
     <p>Segue ranking atualizado das top 25 lojas que tem maior faturamento</p>
     <p>{ranking}</p>
     """

     msg = email.message.Message()
     msg['Subject']= 'Relatório semanal'
     msg['From']= 'paloma.moozer@gmail.com'
     msg['To']= 'paloma.moozer@gmail.com'
     password = 'chrrietkbuchjxhm'
     msg.add_header('Content-Type','text/html')

     msg.set_payload(corpo_email)

     s= smtplib.SMTP('smtp.gmail.com: 587')
     s.starttls()

     s.login(msg['From'],password)
     s.sendmail(msg['From'],[msg['To']], msg.as_string().encode('utf-8'))
     print('Email enviado')


enviar_email()