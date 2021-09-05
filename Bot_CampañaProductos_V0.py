##Bot_CampañaProductos_V0

#Instalar el modulo gspread 

#Librerias de python que vamos a estar usando
import gspread
import os 
from datetime import datetime
import smtplib
from email.message import EmailMessage
from datetime import timedelta  

#Cambia el directorio en donde nos encontramos
os.chdir("D:\matias\Python scrapping\Bot gmail\Bot marketing") #Poner entre comillas la direccion donde tenemos todos los archivos
#Nuestros datos de mail
mail_remitente = os.environ.get("GUSER") #Esta es una variable de entorno con nuestro mail por un tema de seguridad, en "" va el nombre que le pusimos a la variable   
mail_remitente_contraseña= os.environ.get("GPASS") #Esta es una variable de entorno con nuestra contraseña de mail por un tema de seguridad, en "" va el nombre que le pusimos a la variable   
#Fecha actual 
fecha=datetime.now().strftime("%d/%m/%Y")

#Esto nos va a obtener el mensaje que tengamos guardado predeterminado en un archivo txt. En caso de poner otro nombre al archivo del mensaje cambiar donde dice cuerpo txt por el nombre correcto
with open("cuerpo.txt",encoding="utf-8") as f:
    cuerpo = f.readlines()

with open("cuerpo2.txt",encoding="utf-8") as f:
    cuerpo2 = f.readlines()

#Todo sobre google sheet
gc = gspread.service_account(filename="credentials.json")#entre comillas poner el nombre del archivo json con nuestras credenciales de google
sh= gc.open_by_key("")#Aca hay que poner la key del google sheet 
worksheet= sh.sheet1
filas = worksheet.get_all_records()
print(filas)



for index, fila in enumerate(filas, start=2):
    if fila['Enviar campaña'].upper() != "NO" :

        if fila['Contacto'].upper() != "NN" or "" :  
            #Conseguimos los datos de la empresa
            cliente_empresa = fila["Empresa"]
            cliente_mail = fila["Mail de contacto"]
            cliente_nombre = fila["Contacto"]                            

            #Vamos a crear nuestro mail
            encabezado= 'encabezadomailplaceholder'
            mensaje='Buenos dias '+ cliente_nombre + ".\n" + "".join(cuerpo)  

            msg = EmailMessage()
            msg['Subject'] = encabezado
            msg['From'] = mail_remitente
            msg['TO'] = cliente_mail 
            msg.set_content(mensaje)

            #TENER CUIDADO PESO MAXIMO DE LA SUMA DE ARCHIVOS, NO PUEDE SUPERAR 25M
            files = ["a.pdf","b.pdf","c.pdf"]
                #Itera y adjunta uno por uno los archivos 
            for file in files:
                with open (file, 'rb') as f:
                    file_data = f.read()
                    file_name = f.name
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
            #Manda el mail. Tener cuidado con que el with se encuentre afuera del for de los archivos ya que una vez probando me mando un mail cada vez que adjuntaba un archivo
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(mail_remitente, mail_remitente_contraseña)
                smtp.send_message(msg)            

            columna=list(fila.keys()).index("Comentarios") + 1 #Utiliza el indice de Comentarios  para saber la columna y el + 1 es por el desfase de indices entre Python y Sheets
            comentario_anterior = str(fila['Comentarios']) #Nuestro comentario  viejo
            comentario_nuevo=comentario_anterior + ".Contactado" +str(fecha)#Nuestro comentario nuevo que va a ser puesto en la celda de Comentarios del 
            worksheet.update_cell(index, columna, comentario_nuevo)  #Actualiza la celda de Comentarios y para saber las coordenadas usa el index como fila y la funcion columna como columna
       
        else:
            cliente_empresa = fila["Empresa"]
            cliente_mail = fila["Mail de contacto"]
                                      

            #Vamos a crear nuestro mail
            encabezado= 'encabezadomailplaceholder'
            mensaje='Buenos dias ' + ".\n" + "".join(cuerpo2)  

            msg = EmailMessage()
            msg['Subject'] = encabezado
            msg['From'] = mail_remitente
            msg['TO'] = cliente_mail 
            msg.set_content(mensaje)

            #TENER CUIDADO PESO MAXIMO DE LA SUMA DE ARCHIVOS, NO PUEDE SUPERAR 25M
            files = ["a","b","c"]
                #Itera y adjunta uno por uno los archivos 
            for file in files:
                with open (file, 'rb') as f:
                    file_data = f.read()
                    file_name = f.name
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
            #Manda el mail. Tener cuidado con que el with se encuentre afuera del for de los archivos ya que una vez probando me mando un mail cada vez que adjuntaba un archivo
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(mail_remitente, mail_remitente_contraseña)
                smtp.send_message(msg)            

            columna=list(fila.keys()).index("Comentarios") + 1 #Utiliza el indice de Comentarios  para saber la columna y el + 1 es por el desfase de indices entre Python y Sheets
            comentario_anterior = str(fila['Comentarios']) #Nuestro comentario  viejo
            comentario_nuevo=comentario_anterior + ".Contactado" +str(fecha)#Nuestro comentario nuevo que va a ser puesto en la celda de Comentarios del 
            worksheet.update_cell(index, columna, comentario_nuevo)
else:
        pass 

print("Finalizado")