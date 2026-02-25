<%
if session("bandera")<>-1 then

Response.Redirect "caduca.asp"

else

	public objEnlaceBDD
	public objDatosBDD

	Set objEnlaceBDD = Server.CreateObject("EnlaceBDD.clsEnlaceBDD")
	Set objDatosBDD = Server.CreateObject("ADODB.Recordset")
	Set objDatosBDD1 = Server.CreateObject("ADODB.Recordset")
	Set objDatosBDD2 = Server.CreateObject("ADODB.Recordset")


	if objEnlaceBDD.Conectar("") then
		'Estado de la conexion: desconectado
		Response.Redirect "noconectado.asp"
	end if

Dim fechactual, codigop, codigo
fechactual = Date

'Ingreso de datos en cardex

usuario_cardex=Request.Form("usuario_cardex")
clave_cardex=Request.Form("clave_cardex")
nombre_usuario=Request.Form("nombre_usuario")
descripcion_usuario=Request.Form("descripcion_usuario")
permiso_usuario=Request.Form("permiso_usuario")
estado_usuario=Request.Form("estado_usuario")
fecha_ingreso=fechactual

objEnlaceBDD.Consultar "USUARIO", "", "USUARIO='" & usuario_cardex & "' ", ""     
Set objDatosBDD = objEnlaceBDD.Resultado


if not objDatosBDD.EOF then %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY>

<P align="center"><b><font color="#FF0000">LOS SIENTO</font></b></P>
<P align="center"><b><font color="#FF0000">EL USUARIO INGRESADO YA EXISTE</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="ingreso_usuario_cardex.asp" target=principal>VUELVALO A INTENTAR</a></font></P>

<P align="center">&nbsp;</P>

</BODY>
</HTML>

<%
objDatosBDD.Close
else


objEnlaceBDD.Consultar "USUARIO", "", "", "COD_USUARIO ASC"     
Set objDatosBDD1 = objEnlaceBDD.Resultado

objDatosBDD1.MOveLast

codigop=objDatosBDD1.Fields(0)

codigo=codigop+1


IF usuario_cardex <> "" THEN
usuario_cardex=usuario_cardex
ELSE
usuario_cardex=""
END IF

IF clave_cardex <> "" THEN
clave_cardex=clave_cardex
ELSE
clave_cardex=""
END IF

IF nombre_usuario <> "" THEN
nombre_usuario=nombre_usuario
ELSE
nombre_usuario=""
END IF

IF descripcion_usuario <> "" THEN
descripcion_usuario=descripcion_usuario
ELSE
descripcion_usuario=""
END IF

IF estado_usuario <> "" THEN
estado_usuario=estado_usuario
ELSE
estado_usuario="Ac"
END IF

IF permiso_usuario <> "" THEN
permiso_usuario=permiso_usuario
ELSE
permiso_usuario="U"
END IF


IF fecha_ingreso <> "" THEN
fecha_ingreso=fecha_ingreso
ELSE
fecha_ingreso="100/01/01"
END IF


Dim datos_cardex(8)

datos_cardex(0)="'"& codigo &"'"
datos_cardex(1)="'"& usuario_cardex &"'"
datos_cardex(2)="'"& clave_cardex &"'"
datos_cardex(3)="'"& nombre_usuario &"'"
datos_cardex(4)="'"& descripcion_usuario &"'"
datos_cardex(5)="'"& estado_usuario &"'"
datos_cardex(6)="'"& permiso_usuario &"'"
datos_cardex(7)="'"& fecha_ingreso &"'"

objEnlaceBDD.AlmacenarSql "usuario", datos_cardex, 8


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P align="center">&nbsp;</P>


<P align="center"><font size="2" color="#0000FF">
SUS DATOS HAN SIDO REGISTRADOS. <a href="ingreso_usuario_cardex.asp" target=principal>VOLVER A INGRESAR</a></font></P>

</BODY>
</HTML>
<%
objDatosBDD.Close
end if 
end if 
%>
