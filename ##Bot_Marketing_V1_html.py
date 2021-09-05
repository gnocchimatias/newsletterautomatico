##Bot_Marketing_V1_html

#Instalar el modulo gspread y tambien el modulo radar 
#PARA CADA CAMPAÑA HAY QUE CAMBIAR EL HTML, 
#Librerias de python que vamos a estar usando
import gspread
import os 
from datetime import datetime
import smtplib
from email.message import EmailMessage
from datetime import timedelta  
import radar
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText   

#Cambia el directorio en donde nos encontramos
os.chdir("D:\matias\Python scrapping\Bot gmail\Bot marketing") #Poner entre comillas la direccion donde tenemos todos los archivos
#Nuestros datos de mail
mail_remitente = os.environ.get("GUSER") #Esta es una variable de entorno con nuestro mail por un tema de seguridad, en "" va el nombre que le pusimos a la variable   
mail_remitente_contraseña= os.environ.get("GPASS") #Esta es una variable de entorno con nuestra contraseña de mail por un tema de seguridad, en "" va el nombre que le pusimos a la variable   


#Fecha actual y fechas que usaremos mas adelante para evitar enviar todos los mails el mismo dia
fecha=datetime.now().strftime("%d/%m/%Y")
dia_start=datetime.now().strftime("%m/%d/%Y")
dia_stop=(datetime.now() + timedelta(days=4))


#Esto nos va a obtener el mensaje que tengamos guardado predeterminado. En caso de poner otro nombre al archivo del mensaje cambiar donde dice cuerpo txt por el nombre correcto
with open("cuerpo.txt",encoding="utf-8") as f:
    cuerpo = f.readlines()

#Todo sobre google sheet
gc = gspread.service_account(filename="credentials.json")#entre comillas poner el nombre del archivo json con nuestras credenciales de google
#ACORDARSE DE QUE SI SE CAMBIA LA PLANILLA SE CAMBIA LA KEY DE ABAJO
sh= gc.open_by_key("")#Entre las comillas va la key de google sheets
worksheet= sh.sheet1
filas = worksheet.get_all_records() #Crea una lista de cada fila con diccionarios y las keys del diccionario van a ser nuestro encabezado.
 


for i, fila in enumerate(filas, start=2):
    
    
    if fila['Fecha a contactar'] == '': 
        fecha_aleatoria_contacto=radar.random_datetime(start=dia_start, stop=dia_stop)#Asigna fecha aleatorias automatica en el rango de 1 semana para todos los contactos q no tengan 1er contacto puesto
        fecha_primer_contacto=fecha_aleatoria_contacto.strftime("%d/%m/%Y")
        
        columna_primer=list(fila.keys()).index('Fecha a contactar') + 1  
                 
        worksheet.update_cell(i, columna_primer, fecha_primer_contacto)  #Actualiza la celda de la fecha del primer contacto
        
        

filas_actualizadas = worksheet.get_all_records()


#Itera por cada fila e index nos va a servir para saber en que numero de fila estamos, hacemos q comienze en 2 por el desfase que hay entre python y spreadsheet 
for index, fila_actualizada in enumerate(filas_actualizadas, start=2):
    

    if fila_actualizada['Fecha a contactar'] == fecha and fila_actualizada['Enviar campaña\n(SI o vacio /  NO)'].upper() != "NO" and fila_actualizada['Contactado'].upper() != "SI"  :  
        #Conseguimos los datos de la empresa
        cliente_empresa = fila_actualizada["Empresa"]
        cliente_nombre = fila_actualizada["Contacto"]
        cliente_mail = fila_actualizada["Mail de contacto"]
                                    

        #Vamos a crear nuestro mail
        encabezado= 'encabezadomailplaceholder'
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = encabezado
        msg['From'] = mail_remitente
        msg['TO'] = cliente_mail
        #Aca va nuestro mensaje HTML
        html = """\
            <html lang="en">

            
            </html>
            """ 
        part1 = MIMEText(html, 'html')



        msg.attach(part1)
        # Envia el mail via smtp.
        mail = smtplib.SMTP('smtp.gmail.com', 587)

        mail.ehlo()

        mail.starttls()

        mail.login(mail_remitente, mail_remitente_contraseña)
        mail.sendmail(mail_remitente, cliente_mail, msg.as_string())
        mail.quit()

        #Actualizar fila para dejar en claro que ya se contacto
        columna_primer=list(fila_actualizada.keys()).index('Contactado') + 1
        worksheet.update_cell(i, columna_primer, "SI")

        
