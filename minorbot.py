import pywhatkit
import pyautogui
import datetime
import csv
import ctypes
import random
import time
from tkinter import *

hwnd = ctypes.windll.kernel32.GetConsoleWindow()
ctypes.windll.kernel32.SetConsoleTitleW('Minorbot WhatsApp Sender versión 1.0.240129')

titulo = """
+----------------------------------------------------------------------------------------------------------+
|                                                                                                          |
|                                                ▄▄▀▀█▀▀▄▄                                                 |
|                                               ▐▄▌ ▀ ▀ ▐▄▌                                                |
|                                                 █ ▄▄▄ █  ▄▄                                              |
|                                                 ▄█▄▄▄█▄ ▐  ▌                                             |
|                                               ▄█▀█████▐▌ ▀ ▐                                             |
|                                               ▀ ▄██▀██▀█▀▄▄▀                                             |
|                                                                                                          |
|                              M I N O R B O T   W H A T S A P P   S E N D E R                             |
|                        Envía mensajes WhatsApp de manera masiva con San Changuito                        |
|                                                                                                          |
|                                          Minorsoft Corporation                                           |
|                                           versión 1.0.240129                                             |
|                                                                                                          |
|                                 Copyright (c) Minorsoft Corporation 2024                                 |
|                                      Todos los Derechos Reservados                                       |
|                                   San Changuito es una marca registrada                                  |
|                                         de Minorsoft Corporation                                         |
|                                                                                                          |
+----------------------------------------------------------------------------------------------------------+

 Bienvenido a nuestro maravilloso "Minorbot WhatsApp Sender", podrás enviar WhatsApps de una manera masiva.

 ¿No tienes idea de como funciona MinorBot? Nosotros tampoco, solo San Changuito lo sabe ¡Dale una banana!.
       Consulta el módulo de Ayuda, lo puedes observar en el menú inferior, tan sólo escribe "Ayuda".

   Explora los módulos disponibles en el menú, elige uno y escríbelo como el módulo que deseas consultar.
                       _____________________________________________________________
                      |                                                             | 
----------------------| MÓDULOS DISPONIBLES PARA QUE SAN CHANGUITO SE PONGA A JALAR |-----------------------
                      |_____________________________________________________________|
"""
print(titulo)

while True:                       
    print('Módulo "Mensajes"   → ¿Listo para changarles la vida a los deudores? ¡Vamos, démosles caña de una vez!')
    print('Módulo "Plantillas" → Visitemos las plantillas en el jardín, ¡Posiblemente ahí ande San Changuito!')
    print('Módulo "Ayuda"      → ¿No funciona nuestro maravilloso Minorbot? ¡Eso es imposible, pero anda revisemos!')
    print('Módulo "Copyright"  → Revisa los derechos reservados, ¡Somos la mera crema desarrollando en Minorsoft!')
    print('Módulo "Creditos"   → ¿Quieres conocer al programador sin vida social? ¡Cómprale una caguama bien fría!')
    print('Módulo "Licencia"   → Realmente no te interesa esto, pero debemos ponerlo ¡Sabemos que adoras el chisme!')
    print('Módulo "Salir"      → Deseas salir de nuestro maravilloso Minorbot, ¡San Changuito te va a extrañar mucho!')
    print('\n')

    modulo = input('Escribe un módulo, para que San Changuito le de caña en corto: ').lower()

    if modulo == 'mensajes':
        texto_enviar = """
        Quieres poner a jalarsela a San Changuito... ¿O cómo se dice cuándo le ponen mucho jale?

        San Changuito se pondrá en ChangaMadriza loca a leer tus registros e intentar abrir Google Chrome para enviar 
        WhatsApps, una vez que lo pongas a jalar no lo interrumpas ni por todas las bananas del mundo ¿De acuerdo?

        En caso de que quieras darle break (detener el envio de WhatsApps) a San Changuito tan solo cierra Minorbot,
        Estás intentando enviar WhatsApps pero no se envían, revisa la información en módulo de Ayuda ¡No seas flojo!
        """
        print(texto_enviar)

        win = Tk()
        screen_width = win.winfo_screenwidth()
        screen_height= win.winfo_screenheight()

        while True:
            continuar = input("¿Quieres que San Changuito le de caña bien macizo? (S/N): ")
            if continuar.lower() == 's':
                print('San Changuito empezará a jalar, calmala, está changandose una banana pa´ agarrar power\n')
                time.sleep(3)

                plantillas = [
                    '¡Buen día estimado(a) {nombre}! Mi nombre es Lic. {representante} Representante Legal de {empresa}, me comunico con respecto al servicio que tenemos con usted debido que tenemos atraso con el mismo con número de crédito {credito}, requerimos que se pueda poner al corriente del pago de {deuda} M.N. espero su respuesta por este medio antes del {fecha}, en caso de no tener respuesta se actuará correspondiente a las clausulas del contrato. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    'El día de hoy {fecha} validamos la dirección que usted {nombre} agrego en el contrato con nuestro representado {empresa}, el cuál esta solicitando el cobro total del adeudo de {deuda} M.N. en el domicilio por incumplimiento del contrato {credito} contamos con su pagaré y contrato firmado, notifique a sus familiares de la llegada de las dependencias de seguridad así como los abogados para realizar la recuperación total en un plazo máximo de 24 horas, notificación conforme a derecho. Lic. {nombre}. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    'Apreciable {nombre} le comentamos que el día de mañana procederemos en su domicilio a partir de medio día, para realizar la recuperación por un total de {deuda} M.N. del adeudo que tiene con {empresa} por incumplimiento de contrato {credito}, el día de hoy {fecha} se notifico a dependencias de seguridad así como a nuestro abogado {representante} que llevará a acabo la diligencia. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    'Estimado(a) {nombre} tiene un pago pendiente de {deuda} M.N. correspondiente a su número de crédito {credito} con {empresa} le invitamos a que realice su pago antes del {fecha}, cualquier duda responda este WhatsApp con el Lic. {representante}. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    '{nombre} le mencionamos que tiene un pendiente de adeudo con {empresa} por la cantidad de {deuda} M.N. del número de crédito {credito}, realice su pago antes del {fecha} para que no le genere inconvenientes financieros. *Lic. {representante}*. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    'Su número de crédito {credito} ha sido destinada a un despacho de abogados para la recuperación total del adeudo {deuda} M.N. por incumplimiento de contrato entre nuestro representado legal {empresa} y el C. {nombre} a la fecha {fecha}. ¿No deseas ver afectado tu Buro de Crédito? Respondeme este WhatsApp, con gusto te ayudo, Lic {representante}. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    '{nombre} no ha realizado su pago de {deuda} M.N. a la fecha {fecha} ante {empresa} por lo cuál se procede a visitar el domicilio por el cobro total del número de crédito {credito}, le saluda el Lic. {representante}. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    '¡Enhorabuena {nombre}! Tenemos un descuento autorizado para liquidar su adeudo de {deuda} M.N. con {empresa} evitando la diligencia en la dirección de su domicilio agregado en el contrato, si le interesa el descuento solicítelo respondiendo a este Whatsapp con la clave B{empresa}{credito}-24 con el Lic. {representante}, descuento solo válido por hoy {fecha} ¡No esperes más {nombre}! Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.',
                    '*¡{nombre} tomaremos su desinterés como negativa de pago en su crédito {credito}!*, hemos dictaminado su contrato con {empresa} el día de hoy {fecha} por tal motivo su contrato será transferido a nuestro despacho de abogados con el Lic. {representante}, para generar la recuperación de {deuda} M.N. en el domicilio, llevaremos los documentos que firmó al inicio del contrato. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.'
                    'Este es un mensaje de prueba por parte de {representante} el día {fecha} para revisar el número de crédito {credito} correspondiente al cliente {empresa} del deudor {nombre} porque tiene un adeudo por {deuda} M.N. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada'
                ]

                random.shuffle(plantillas)
                
                with open('registros.csv', 'r', encoding='utf-8-sig') as registro:

                    lector_registros = csv.reader(registro)
                    next(lector_registros, None)
                    next(lector_registros, None)

                    for row in lector_registros:

                        if any(cell.strip() for cell in row):

                            tiempo_de_espera = random.uniform(3, 10)
                            print(f"San Changuito enviara el mensaje, en {round(tiempo_de_espera)} segundos...\n")
                            time.sleep(tiempo_de_espera)

                            plantilla_seleccionada = random.choice(plantillas)
                            plantilla_interpolada = plantilla_seleccionada.format(
                                nombre=row[1], representante=row[6], empresa=row[2],
                                credito=row[4], deuda=row[3], fecha=row[5], telefono=row[7]
                            )

                            time_now = datetime.datetime.now()
                            time_to_send = time_now + datetime.timedelta(minutes=1)

                            hours = time_to_send.hour
                            minutes = time_to_send.minute

                            pywhatkit.sendwhatmsg(f"+52{row[0]}", plantilla_interpolada, hours, minutes)

                            pyautogui.moveTo(screen_width / 2, screen_height / 2)

                            # time.sleep(20)

                            pyautogui.click()
                            pyautogui.press('enter')

                            # time.sleep(5)
                            # pyautogui.hotkey('ctrl', 'w')

                            print('\nMensaje enviado correctamente por San Changuito, ¡Se la rifo chidito! ¿No?...\n')

                input('Presiona "Enter" para salir, ya termino San Changuito bien changado de tanto jale...')
                break
            elif continuar.lower() == 'n':
                print('San Changuito se irá de año sabático pero antes te regresara al menú de módulos...')
                time.sleep(3)
                break
            else:
                texto_else = """
                Changa..da madre, opción no válida. Por favor no hagas enojar a San Changuito.
                Selecciona "s" para continuar o "n" para de este módulo...\n\n
                """
                print(texto_else)
    elif modulo == 'plantillas':
        texto_plantillas = """
            Te mostramos las plantillas (No olvides regarlas continuamente) que se estarán enviando en cada 
            mensaje a los deudores lacras ¡Esperamos te paguen mucho dinero!...

            PLANTILLA 1:
            ¡Buen día estimado(a) {DEUDOR}! Mi nombre es Lic. {EMPLEADO} Representante Legal de {PRODUCTO}, 
            me comunico con respecto al servicio que tenemos con usted debido que tenemos atraso con el mismo 
            con número de crédito {CREDITO}, requerimos que se pueda poner al corriente del pago de {DEUDA} M.N. 
            espero su respuesta por este medio antes del {FECHA}, en caso de no tener respuesta se actuará 
            correspondiente a las clausulas del contrato. Contesta este WhatsApp o manda uno a {TELEFONO} 
            para atención personalizada.
            
            PLANTILLA 2:
            El día de hoy {FECHA} validamos la dirección que usted {DEUDOR} agrego en el contrato con nuestro 
            representado {PRODUCTO}, el cuál esta solicitando el cobro total del adeudo de {DEUDA} M.N. en el 
            domicilio por incumplimiento del contrato {CREDITO} contamos con su pagaré y contrato firmado, 
            notifique a sus familiares de la llegada de las dependencias de seguridad así como los abogados 
            para realizar la recuperación total en un plazo máximo de 24 horas, notificación conforme a derecho. 
            Lic. {EMPLEADO}. Contesta este WhatsApp o manda uno a {TELEFONO} para atención personalizada.

            PLANTILLA 3:
            Apreciable {nombre} le comentamos que el día de mañana procederemos en su domicilio a partir de medio 
            día, para realizar la recuperación por un total de {deuda} M.N. del adeudo que tiene con {empresa} 
            por incumplimiento de contrato {credito}, el día de hoy {fecha} se notifico a dependencias de 
            seguridad así como a nuestro abogado {representante} que llevará a acabo la diligencia. Contesta este 
            WhatsApp o manda uno a {telefono} para atención personalizada.

            PLANTILLA 4:
            Estimado(a) {nombre} tiene un pago pendiente de {deuda} M.N. correspondiente a su número de crédito 
            {credito} con {empresa} le invitamos a que realice su pago antes del {fecha}, cualquier duda responda 
            este WhatsApp con el Lic. {representante}. Contesta este WhatsApp o manda uno a {telefono} para 
            atención personalizada.

            PLANTILLA 5:
            {nombre} le mencionamos que tiene un pendiente de adeudo con {empresa} por la cantidad de {deuda} M.N. 
            del número de crédito {credito}, realice su pago antes del {fecha} para que no le genere inconvenientes 
            financieros. *Lic. {representante}*. Contesta este WhatsApp o manda uno a {telefono} para atención 
            personalizada.

            PLANTILLA 6:
            Su número de crédito {credito} ha sido destinada a un despacho de abogados para la recuperación total 
            del adeudo {deuda} M.N. por incumplimiento de contrato entre nuestro representado legal {empresa} y el 
            C. {nombre} a la fecha {fecha}. ¿No deseas ver afectado tu Buro de Crédito? Respondeme este WhatsApp, 
            con gusto te ayudo, Lic {representante}. Contesta este WhatsApp o manda uno a {telefono} para atención 
            personalizada.

            PLANTILLA 7:
            {nombre} no ha realizado su pago de {deuda} M.N. a la fecha {fecha} ante {empresa} por lo cuál se procede 
            a visitar el domicilio por el cobro total del número de crédito {credito}, le saluda el Lic. {representante}. 
            Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.

            PLANTILLA 8:
            ¡Enhorabuena {nombre}! Tenemos un descuento autorizado para liquidar su adeudo de {deuda} M.N. con {empresa} 
            evitando la diligencia en la dirección de su domicilio agregado en el contrato, si le interesa el descuento 
            solicítelo respondiendo a este Whatsapp con la clave B{empresa}{credito}-24 con el Lic. {representante}, 
            descuento solo válido por hoy {fecha} ¡No esperes más {nombre}! Contesta este WhatsApp o manda uno a 
            {telefono} para atención personalizada.

            PLANTILLA 9:
            ¡{nombre} tomaremos su desinterés como negativa de pago en su crédito {credito}!, hemos dictaminado su 
            contrato con {empresa} el día de hoy {fecha} por tal motivo su contrato será transferido a nuestro 
            despacho de abogados con el Lic. {representante}, para generar la recuperación de {deuda} M.N. en el 
            domicilio, llevaremos los documentos que firmó al inicio del contrato. Contesta este WhatsApp o manda 
            uno a {telefono} para atención personalizada.

            PLANTILLA 10:
            Este es un mensaje de prueba por parte de {representante} el día {fecha} para revisar el número de 
            crédito {credito} correspondiente al cliente {empresa} del deudor {nombre} porque tiene un adeudo por 
            {deuda} M.N. Contesta este WhatsApp o manda uno a {telefono} para atención personalizada.

        """
        print(texto_plantillas)
    elif modulo == 'ayuda':
        texto_ayuda = """
            Te damos la más cordial bienvenida a Minorbot WhatsApp Sender, una herramienta majestuosa para 
            tus necesidades de mensajería WhatsApp. ¿Es tu primera experiencia con nuestro sistema 
            extraordinario? No te preocupes, presta atención a la siguiente información nuestro maravilloso 
            Minorbot no debería de fallar, o eso queremos creer.

            Ingrese el nombre de cualquier módulo, para acceder al mismo, principalmente al módulo de 
            "enviar" que es el que más te puede servir, básicamente es el motivo principal por el cuál 
            desarrollamos este sistema en Minorsoft Corporation.
            Cada módulo tiene su descripción en el menú inicial, donde se te muestran todos los módulos con 
            los que puedes interactuar ¡Cómo se te pegue la gana, no nos hacemos responsables si explota
            tu computadora o se traba!.

            ¡Ahora estás en manos de programadores...
            de los mejores en Minorsoft Corporation, eres un condenado con mucha suerte!

            ¿Cómo puedes empezar a hacer envío masivo de mensajes de texto por WhatsApp?

            Para poder enviar mensajes de WhatsApp en chingamadriza a los deudores primero tienes que tener 
            preparado tu archivo Excel CSV (delimitado por comas), este archivo tiene que tener 7 columnas... 
            (Podemos llamarles las 7 keywords de la cobranza, aunque ni cobres ¡Pero bueno...!)

            ¿Cuáles son esas columnas y como deben ir ordenadas? ¿Tienen que tener encabezado? 
            ¿Qué formato deben llevar? ¿Cuántos registros puedes cargar para cada envío?

            ¡COLUMNAS QUE DEBEN ESTAR EN EL ARCHIVO CSV!
            (SI NO ESTÁN, NO ESPERES QUE FUNCIONE EL SISTEMA, SOMOS PROGRAMADORES... ¡TAMPOCO SON ENCHILADAS!)

            1.- WHATSAPP (Número sin decimales) - ¿A quién molestaremos con el WhatsApp? 
            ¡No somos adivinos, colócalo! Y no, no se pone el "+52" Solo leemos número de 10 dígitos.

            2.- DEUDOR (Texto) - Es el nombre del deudor, sí... ¡La rata mugrosa que nos debe mucho dinero!

            3.- PRODUCTO (Texto) - Es el nombre del cliente corporativo al que pertenece la ratota que nos 
            debe dinero, perdón nos explayamos, era "Cliente comercial" alias el deudor.

            4.- DEUDA (Moneda) - Es lo que nos debe la rata de dos patas, entonces recuerda que entre más 
            cobres ¡Tu cheque a fin de mes es más gordo! ¡Aquí lo que sobra es varo!

            5.- CREDITO (Número sin decimales) - ¿Cómo sabes a quién ver**s le cobras?, exacto con el número 
            de subcrédito, no tampoco... aquí no aplica  el número de subcrédito, 
            ¿Quién ocupa ese tonto dato?... ¡Sólo el de sistemas!

            6.- FECHA (Fecha en letra ¿Sabes Excel?) - ¿Qué día tiene como límite la rata para pagar? Bueno, 
            eso es lo que va en esta columna.

            7.- EMPLEADO (Texto) - La pura crema y nada de salsa, ¡Esa persona eres Tú!, o al menos tu 
            nombre en este mundo de la cobranza.

            8.- TELEFONO (Número sin decimales) - ¡No te gusta que te roben tus pagos! Amarralos, haz que 
            las ratas hables contigo directamente...¿Espera qué, las ratas hablan?

            Acabas de ver el orden como deben ir las columnas, ¡No ocupamos encabezado, así que ni lo 
            coloques! También te acabamos de decir el formato que deben tener para que no hagas más 
            preguntas.
            No cargues más de 80 registros en el archivo Excel CSV delimitado por comas, tampoco exageres.
            ¡No quieras que hagamos tu trabajo, suficiente tuvimos con el desarrollo de MinorBot!

            ¿Qué sucede o que debes hacer en caso de que no funcione nuestro querido MinorBot?
            Opinamos seriamente que te cambies de call center ¡Aquí no te queremos! No te creas, aquí te 
            damos unos tips, solo síguelos al pie de la letra y no molestes al de sistemas ¿Ok?

            1.- Reza 10 Padres Nuestros, y avientate una persignada ¡Oremos!
            2.- Preferentemente utiliza Google Chrome para usar WhatsApp Web (No desktop)
            3.- Utiliza la combinación de CTRL + SHIFT + R para eliminar el caché en el navegador (Chrome)
            4.- Reinicia tu computadora
            5.- No funciona el envío de WhatsApps ¿Revisaste que el archivo este correcto?
            6.- Ya de plano si no funciona ¡Mejor cámbiate de empleo o al menos de call center!

            Por nada del mundo, y hablamos en serio...
            Por nada del mundo mundial debe de faltar el archivo CSV delimitado por comas a nivel de
            nuestro poderosísimo MinorBot, en la carpeta donde está MinorBot, o nuestro sistema se irá 
            de año sabático ¡No bromeamos!

            ¿ENTONCES QUE DICES, NOS INVITAS UNA CAGUAMA? ¡Gracias por leer nuestra biblia, fomentemos la lectura!
        """
        print(texto_ayuda)
    elif modulo == 'copyright':
        texto_copyright = """
            Copyright (c) 2024 Minorsoft Corporation.
            Todos los Derechos Reservados.

            Copyright (c) Buró Jurídico Integral de Cobranza S.C.
            Todos los Derechos Reservados.
        """
        print(texto_copyright)
    elif modulo == 'creditos':
        texto_creditos = """
            Agradecemos profundamente a Buró Jurídico Integral de Cobranza S.C. y a un selecto 
            grupo de individuos por su invaluable respaldo al avance y evolución de MinorBot.

            MinorBot, una creación cuidadosamente elaborada con dedicación y profesionalismo 
            por Minorsoft Corporation, bajo los derechos de autor (c) 2024. El desarrollo técnico 
            de esta innovadora plataforma ha sido desarrollado por Ing. Juan Morales Minor.

            ¡QUE DIOS NOS AMPARE. HEMOS CAÍDO EN MANOS DE INGENIEROS!

            Si estás interesado en contribuir a la mejora continua de este sistema o tienes 
            sugerencias para perfeccionarlo aún más, te invitamos cordialmente a ponerte en 
            contacto con nosotros:

            Correo Electrónico: jmoralesminor@gmail.com

            ¡Apóyanos con la adquisición de una deliciosa caguama bien fría para seguir mejorando MinorBot!
        """
        print(texto_creditos)
    elif modulo == 'salir':
        print('No te vamos a extrañar para nada, ¡Es broma... pero si quieres no es broma!')
        time.sleep(3)
        break
    elif modulo == 'SanChanguito':
        texto_ee = """
         ▄▄▀▀█▀▀▄▄        ▄▄▀▀█▀▀▄▄        ▄▄▀▀█▀▀▄▄        ▄▄▀▀█▀▀▄▄        ▄▄▀▀█▀▀▄▄    
        ▐▄▌ ▀ ▀ ▐▄▌      ▐▄▌ ▀ ▀ ▐▄▌      ▐▄▌ ▀ ▀ ▐▄▌      ▐▄▌ ▀ ▀ ▐▄▌      ▐▄▌ ▀ ▀ ▐▄▌    
          █ ▄▄▄ █  ▄▄      █ ▄▄▄ █  ▄▄      █ ▄▄▄ █  ▄▄      █ ▄▄▄ █  ▄▄      █ ▄▄▄ █  ▄▄  
          ▄█▄▄▄█▄ ▐  ▌     ▄█▄▄▄█▄ ▐  ▌     ▄█▄▄▄█▄ ▐  ▌     ▄█▄▄▄█▄ ▐  ▌     ▄█▄▄▄█▄ ▐  ▌
        ▄█▀█████▐▌ ▀ ▐   ▄█▀█████▐▌ ▀ ▐   ▄█▀█████▐▌ ▀ ▐   ▄█▀█████▐▌ ▀ ▐   ▄█▀█████▐▌ ▀ ▐
        ▀ ▄██▀██▀█▀▄▄▀   ▀ ▄██▀██▀█▀▄▄▀   ▀ ▄██▀██▀█▀▄▄▀   ▀ ▄██▀██▀█▀▄▄▀   ▀ ▄██▀██▀█▀▄▄▀

        SAN CHANGUITO SAN CHANGUITO SAN CHANGUITO SAN CHANGUITO SAN CHANGUITO SAN CHANGUITO
        """
        print(texto_ee)
    else:
        print('¡Changa..da madre, ese módulo no es válido! Por favor, escribe un módulo válido...\n\n')