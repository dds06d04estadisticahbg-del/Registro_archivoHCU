<%
if session("bandera")<>-1 then

Response.Redirect "caduca.asp"

else

	public objEnlaceBDD
	public objDatosBDD

	Set objEnlaceBDD = Server.CreateObject("EnlaceBDD.clsEnlaceBDD")
	Set objDatosBDD = Server.CreateObject("ADODB.Recordset")


	if objEnlaceBDD.Conectar("") then
		'Estado de la conexion: desconectado
		Response.Redirect "noconectado.asp"
	end if

Dim fechactual
fechactual = Date

'Modificación de datos en USUARIO

usuario_cardex=Request.Form("usuario_cardex")
clave_cardex=Request.Form("clave_cardex")
nombre_usuario=Request.Form("nombre_usuario")
descripcion_usuario=Request.Form("descripcion_usuario")
permiso_usuario=Request.Form("permiso_usuario")
estado_usuario=Request.Form("estado_usuario")
fecha_ingreso=fechactual

objEnlaceBDD.Consultar "USUARIO", "", "USUARIO='" & usuario_cardex & "' ", ""     
Set objDatosBDD = objEnlaceBDD.Resultado


numero=objDatosBDD.RecordCount

if numero>1 then %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY>

<P align="center"><b><font color="#FF0000">LO SIENTO</font></b></P>
<P align="center"><b><font color="#FF0000">EL USUARIO INGRESADO YA EXISTE</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="modificacion_usuario_cardex.asp" target=principal>VUELVALO A INTENTAR</a></font></P>

<P align="center">&nbsp;</P>

</BODY>
</HTML>

<%

else

cod_usuario=objDatosBDD.Fields(0)

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

objDatosBDD.close

objEnlaceBDD.ConsultarSQL "UPDATE USUARIO SET USUARIO='"&usuario_cardex&"',CLAVE='"&clave_cardex&"',NOMBRE='"&nombre_usuario&"',DESCRIPCION='"&descripcion_usuario&"',ESTADO='"&estado_usuario&"',PERMISO='"&permiso_usuario&"' where COD_USUARIO = "&cod_usuario&" "


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
SUS DATOS HAN SIDO MODIFICADOS. <a href="modificacion_usuario_cardex.asp" target=principal>MODIFICAR OTRO USUARIO</a></font></P>

</BODY>
</HTML>
<%
end if 
end if 
%>
