Attribute VB_Name = "HISTORY"
'######################## 11 / 10 / 2207 #################
'* /REVIVIR: Le puse que si esta VIVO no lo reviva, lo dejaba en bolas antes -.-
'* /Setdesc  Nuevo comando Clickeas al usuario mandas un /setdesc y le cambia la desc...
'######################## 10 / 10 / 2007 #################
' */VERFPS Nº    pones eso y te saltan todos lso FPS con ese numero, no sabia como hacer el else despues fijate... beso :$
' */LEER CLAN  pones eso y lee elc lan
' * Agreegue un nuevo SEND "TODIOSESYCLAN" para el Leer clan
' * /divorciarse
' * /CASAR
' * Puse el sistema de casamiento pero lo deje comentado para optimisarlo
'########## 04 / 10 / 2007 ##########
'* Cambie el /log/numusus.log por /usuarios/numusus.log ( Para al web)
'* Cambie el /LOGIN WESA, Le puse /LOGIN SOFITEAMO  (Login wesa me lo vieron, lo usaron y banie a 15 pj's) :P
'* Puse el /CI   =  a /CT pero con OBJETO, crea un teleport con objeto solo para Privilegios = 4
'######## 2 / 10 / 2007  #######
'* Se creo un nuevo sistema buscador de cheat (Scanner) Solo para DIOSES - ADMIN.
'Metodo de uso.
'/RESETP (SIEMPRE ANTES DE ARRANCAR UN SCAN)
'/AP PTS ( CON ESTO VA A AGREGAR LA PALABRA "PTS" osea que va a revisar los procesos de todos buscando el PTS.
'/CHECKP (ARRANCA LA COMPROBACION)
'/FINALIZARCOMPROBACION ( TERMINA LA COMPROBACION SIEMPRE ESPERAR 15 o 20 SEG despues del /CHECKP)
'/RESETP ( SIEMPRE PONERLO AL FINALIZAR EaL SCAN, SI NO LO PONEN EL /PROCESOS DEJA DE FUNCIONAR )

'* Se modificaron los permisos para " /FORCEWAV /FORCEMIDI para Dioses, tambien el SCanner fue habilitado)
'* Se arreglo el SUBESED = 3 esto es el hechi del bardo que sAca, comida, agua y sta.. ahora funciona bien..
'* Se saco el Toall de /NOMBRE, ahora solo dice a los Admins y al usuario destino.
'* Se reparo el /SMSG (MENSAJE DE SISTEMA) Este abre un FORM en el cliente con el rdata ( mensaje )
'* Se modificaron los colores de varias cosas para mejorar y que no sea TAN Circo...
'* CLIENTE: Ahora en el cliente llega el /PRIVADO ( EL MENSAJE ) en DrawText llega en la pantalla de vision.(PROXVER)
'* Se modifico el /privado ahora lo manda al drawtext del cliente, esta activo pero hasta lap roxima version (Nuevo cliente) no lo van a poder ver ya que necesita si o si el nuevo parcheo...
'* Recompensa nivel 18 todas clases magicas (PETRIFICAR)
': Paraliza a un NPC durante varios segundos usando 250 de mana.
'* /LIMPIARMAPAS
'INFO:Limpia todos los mapas borrando todos los objetos que se encuentren en el piso, el mismo tiene de demora 1 minuto antes del limpiado.
'* /NOCHESI

'INFO: Activa un efecto noche en el cliente.

'* /NOCHENO

'INFO: Desactiva el efecto noche.

'* /PROCESOS NICK

'INFO: Revisa los procesos del usuario.

'* /CERRARPROCESO NICK@NUMERO

'INFO: Cierra el numero de proceso ya visto en /procesos
'* /PRIVADO NICK@MENSAJE

'INFO: Manda un mensaje a la consola del usuario seleccionado en "NICK".
'* /ENCUESTA PREGUNTA
'INFO: Manda una encuesta global, los usuarios deberan responder con /SI o /NO segun la pregunta situada en "PREGUNTA"

'* /CERRAR
'INFO: Cierra la encuesta ya creada con /ENCUESTA.

'* /RETOACTIVADO
'INFO: Desactiva/Activa los retos (1vs1)

'* /PAREJASACTIVADA
'INFO: Desactiva/Activa los retos (2vs2)

'* /OFRECER NICK CANTIDAD
'INFO: Oferta por la cabeza de un X usuario por x $$$ (OFERTAS)

'* /OFERTASACTIVADAS
'INFO: Desactiva/Activa las ofertas.

'* /LIMPIAROFERTAS
'INFO: Resetea todas las ofertas hasta el momento.


'* /SOPORTEACTIVADO
'INFO: Desactiva/Activa el soporte inmediato.

'* /CHECKFPS

'INFO: Checkea a todos buscando usuarios con un menor de 5FPS.

'* /FPS NICK
'INFO: Checkea los FPS de un usuario determinado en "NICK".

'* /CHEATALL
'INFO: Checkea a todos los usurios con posible cliente externo. Ojo con esto puede que el usuario este "LOGGEADO" y no con externo, si tiene externo al momento de usar /cheatall AUTOMATICAMENTE el servidor lo expulsa del juego a la persona.

'* /CHEATCLICK
'INFO: Se selecciona a un usuario (PREVIO CLICK) y se escribe el comando /Cheatclick este manda un mensaje a nuestra consola diciendo si el usuario esta CON O SIN CLIENTE OFFICIAL!.

'* /PANELGM
'INFO: Se implemento un completo panel GM para los mas Newbies. =) :P

'* /CONSEJO
' 1= Ciuda 0=Crimi - 1=Activo 2=OFF
' /CONSEJO NICK@1@1    Nick es consejo ciudadno activo.
'INFO: No explico porque es solo para la direccion :P pero bueno les doy una idea general, desde el juego ahora se pueden aplicar los consejos y Concilios de dichas facciones.

'*/LASTEMAIL NICK
'INFO: Checkea el email de un user estando ONLINE/OFFLINE (SOLO DIRECCION)

'* /CPASS PJSINPASS@PJCONPASS
'INFO: Cambia la password de un usuario (SOLO DIRECCION)


'* /CEMAIL NICK-nuevomail
'INFO: Cambia el email de un usuario (SOLO DIRECCION)

'* /DATS
'INFO: Actualiza los dats del servidor. (OBJ)

'* /DATSFULL
'INFO: Actualiza por completo los DAT's del servidor (ALL)

'* /SILENCIAR MINUTOS NICK
'INFO: Silencia a un usuario por X minutos (TODOS LOS RANGOS)

'* /PING
'INFO: Comando para saber el ping entre cliente servidor lo normal seria entre 63 y 133.

'* Sistema de torneos automaticos.
'INFO: Torneos automaticos, por nivel, clase, level, cupos maximos.

'* Rebalanceo completo del servidor
'INFO: Se realizo un nuevo balance, en base a jugabilidad en parejas y no invididuales.

'* Consola de clanes
'INFO: Se activa desde el menú opciones y permite una mejor visualidad de los mensajes de una guilds.

'* /SOPORTE
'INFO: El usuario puede mandar un mensaje al acto, el mismo llegara a la consola de lo Administradores (ALL PRIVILEGIOS)

'* Pantala FULLSCREN.
'INFO: Podemos usar F9 y veremos una pantalla mas grande, con un rango mayor de visión permitiendo usar este sistema para Queso y eventos o para fotografías.


'* Sistema de transparencias
'INFO: Contamos con un sistema de transparencia que permite a los usuarios tener una visión agradable del juego en si.

'* Tecla backup (K)
'INFO:Se implemento una pequeña ayuda para los guilds ya que no tenían sentido alguno solo su nombre, ahora apretando la letra “K” los miembros de dichas guilds reciben un mensaje de ayuda con su estado de vida y posición actual.

'* Se agrego una nueva meditacion a NIVEL 43
'INFO:Se agrego una nueva meditación para aquellos que aproximan al nivel 43.

'12 / 9 / 2007

'*Sistema de conquista de castillo
'los usuarios van matan al rey y conquistan al conquistar tiene el privilegio de ir y venir cuadnoquiera /ulla /ircastillo1

'* Se agrego un nuevo anticheat Total
'INFO:Este revisa todos los posibles uso de cheat de todos los usuarios devolviendo en consola todo lo que tienen.

'* Se modifico el portal luminoso
'INFO: Ahora lanzas y aparece un mini efecto de 5 segudnso y despues aparece el portal luminoso.

'* Se agrego un nuevo hechizo (Materializa COMIDA - AGuA)
'INFO:Se puede tirar en zona insegura y cuesta 500 de mana y se venden 500k el hechi, los tiras al terreno y hace un morfi y agua dependiendo el hechi q se tire.

'*Se creo un nuevo sistem a de clanes ( ONda Charfiles) pero para los clanes
'INFOE este modo esperemos que no buguee mas y no cuelga mas el servidor.





















' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'                              POCHOOO
'                              POCHOOO
'                              POCHOOO
'                              POCHOOO
'                              POCHOOO

'26/10/08 > SE AGREGO EL SISTEMA DE /LOGPENA VER SUB LOGPENA
'         > /VERPENAS NICK > ON Y OFFLINE
'          ------------------------------------------------------------------------------
'          ------------------------------------------------------------------------------
'27/11/08 > CLICKEES EN LA CABEZA O EN EL PIE AL BOVEDERO/RESU O ETC
'           Y TE LO TOMA BIEN
'27/11/08 > SE AGREGA EN EL CLIENTE QUE CON LA F SE PAREN LOS FX(CONSOLA AVISA)
'27/11/08 > GM CAMINA SOBRE AGUA
'27/11/08 > SE AGREGA EL NUEVO SISTEMA DE BOVEDA, 3 SEG PARADO Y PUMBA(?
'           KBAO-TRIGGER 10 ;) DORMISTE ROO. AGREGUE TMB QUE SI YA SE
'           ABRIO LA BOVEDA NO SE VUELVA A ABRIR HASTA QUE CAMINE DE NUEVO
'           PORQUE SI NO NO SE LLEGA A EQUIPAR Y ES MOLESTO
'27/11/08 > SE AGREGA EL NUEVO SISTEMA DE SACERDOTE, TRIGGER 11
'27/11/08 > TRIGGER 11 NO PISAN LOS NPCS
'27/11/08 > REMARCO MENOS DE 100K CADA OBJ QUE PONE EL USER EN VENDEDOR
'27/11/08 > REMARCO EL ESTUPIDEZ. NO ES + RANDOM, ES ISQ A DERECHA ETC.
'27/11/08 > ARREGLADO EL BUG DE /ECHARFACCION DE LOS CONSEJOS QUE CUANDO
'           ECHABAN A ALGUIEN ESTANDO OFFLINE Y CON CLAN NO SE BORRABA EL CLAN
'27/11/08 > ARREGLADO EL BUG DEL MINERO QUE SIN LA RECOM PODIA USAR EL PICA FUERTE
'27/11/08 > ARREGLADO EL BUG DE QUE SI TENIAS LA BOVE ABIERTA Y TE PISABAN SE BUGIABA
'          ------------------------------------------------------------------------------
'28/11/08 > ARREGLADO EL BUG DE ECHAR MIEMBROS DEL CLAN OFFLINE
'28/11/08 > ARREGLADO EL BUG DEL /ONLINE EN EL PANELGM EN EL CLIENTE
'28/11/08 > AGREGADO EL MINI MAPA AL CLIENTE * Transparencias. Dble click se cierra.
'28/11/08 > SE AGREGA EL SISTEMA DE CRONOMETRO PARA EL CLIENTE (FRMCRON)
'28/11/08 > SE AGREGA LA FUNC. UNZIP AL CLIENTE Y DEPASO SE COMPRIMEN LOS MINIMAPS.
'           (bug de password, no puedo ponerle password :S)
'28/11/08 > SE AGREGA EL SISTEMA DE FUNC, UNZIP PARA LOS GRAFS DEL CLIENTE!
'28/11/08 > se elimina el log del invisible del GM
'28/11/08 > SE ARREGLO EL IGNORAR. SI ESTA OFFLINE O ES UN GM NO DICE NADA.
'28/11/08 > SE ARREGLO EL BUG DE LA ESTU SI TE MATAN ESTU O DESLOGIAS SE QEDABA
'28/11/08 > SE AGREGO EN EL CLIENTE LOS FPS A 19
'          ------------------------------------------------------------------------------
'30/11/08 > CUANDO LOS SOPORTES NO ESTAN HABILITADOS, ENVIA CARTEL DE AVISO AL USER.
'30/11/08 > SE ACOMODA TODA LA INTERFACE.
'          ------------------------------------------------------------------------------
''3/12/08 > SE AGREGA EL RANK PARA LOS RETOS
''3/12/08 > SE AGREGA LA PIRAMIDE, FALTA TESTEAR :O
''3/12/08 > LADRON NO ROBA COSAS CON REAL = 1 O CAOS = 1
''3/12/08 > PRIMERA PARTE DEL CAPTURE THE FLAG
''          ------------------------------------------------------------------------------
''8/12/08 > SISTEMA DE ORO TRANSFERIBLE. senddata TRSANF(USER@CANT) - Parte Sv
''8/12/08 > SISTEMA DE ORO TRANSFERIBLE. Cliente Panel
''9/12/08 > ARREGLADO EL BUG DEL COMERCIANTE. CAMBIE EN "UserDaObjVenta", COMENTE UNAS COSAS
''9/12/08 > AGREGADO EL COMANDO /MODOQUESTMAP.TELEPORTA USERS CUANDO MUEREN EN X MAP
'          ------------------------------------------------------------------------------
'11/12/08 > NO VALE AUTO MOVERSEEEEEEEEE. /MOVER
'11/12/08 > /MAPKILL MATA A TODO EL MAPA MENOS ADMINS Y GMS
'11/12/08 > SE ARREGLA EL BANT DEFINITIVAMENTE
'          ------------------------------------------------------------------------------
'17/12/08 > AGREGADO EL /VERP PARA VER PALABRAS DEL AP Y ESO.
'          ------------------------------------------------------------------------------
'18/12/08 > CAMBIE LA RECOM DEL PIRATA. NO MAS NOINMO-AHORA ES 5 SEGS
'          ------------------------------------------------------------------------------
'19/12/08 > CAMBIE LA RECOM DEL BARDO. PARA QUE SURGA EFECTO EL APOCA NECESITA LAUD MAGICO
'           EL APU NO ES MAS UNA RECOM, AHORA TIENE RM SIEMPRE Q TENGA LAUD!!!
'19/12/08 > AGREGUE EL BAN PC :) AHORA CORROBORA DEL ARCHIVO BANPC.DAT /BANPC USER NO LO ECHA
'19/12/08 > RECOM NUEVA DEL DRUIDA A LA MITAD
'          ------------------------------------------------------------------------------
'20/12/08 > ACOMODE LA INTERFACE DEL CLIENTE
'20/12/08 > SE ARREGLA EL EDITOR DE MAPAS :)
'          ------------------------------------------------------------------------------
'21/12/08 > AGREGUE PARA QUE HAYA FX CUANDO SE REVIVE Y SE SUBE DE NIVEL
'17/2/09  > AGREGADA LAS BARRAS DE FUERZA Y CELERIDAD
'23/2/09  > SE AGREGAN LAS BARRAS DE EXP, Y SE ACOMODAN LAS DE FUERZA Y CELE
'24/2/09  > AGREGADO EL SISTEMA DE SOPORTE.
'27/2/09  > AGREGADO EL SISTEMA DE CARCEL POR HIERRO Y NO POR TIEMPO.
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'27/2/09  > AGREGADO EL COMANDO /RECOM PARA VER LAS RECOM QUE HAY
'27/2/09  > AGREGADAS LAS RECOMPENSAS DE DRUIDA MAGICO DRUIDA GUERRERO Y DOMADOR
'         > /PROXYC PARA CHCKEAR QUE NO USEN PROXY

'         > SE AGREGA EL /BANTS PARA CORRER TODOS LOS BANT HASTA EL DIA DE HOY
'         > /PRECIOP PARA SETIAR EL PRECIO DEL PARTICIPAR
'           SE AGREGA EL NUEVO SISTEMA DE PAGOS PARA EL CAPTURE THE FLAG



'HACER QUE CON F1 SE VEAN A LOS DEL CLAN DE OTRO COLOR ENCRUZADAS
'PONER PARA BORRAR SOPORTE EN EL CLIENTE


'ENCRIPTAR LOS INDEEEEEX
'HACER EL VIGILADOS- http://furiusao.com.ar/FURIUSAOSTAFF/showthread.php?t=587&page=2


' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
