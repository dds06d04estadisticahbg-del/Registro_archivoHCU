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

Dim fechactual
fechactual = Date

'Modificación de datos en cardex

historia_clinica=Request.Form("historia_clinica")
apellido_paterno=Request.Form("apellido_paterno")
apellido_materno=Request.Form("apellido_materno")
nombres=Request.Form("nombres")
nombre_padre=Request.Form("nombre_padre")
nombre_madre=Request.Form("nombre_madre")
fecha_nacimiento=Request.Form("fecha_nacimiento")
cedula=Request.Form("cedula")
cod_usuario=session("cod_usuario")
fecha_ingreso=fechactual
estado=Request.Form("estado")
sexo=Request.Form("sexo")
localidad=Request.Form("localidad")


objEnlaceBDD.Consultar "CARDEX", "NUM_HISTORIA_CLINICA", "NUM_HISTORIA_CLINICA='" & historia_clinica & "' ", ""   
Set objDatosBDD = objEnlaceBDD.Resultado

numero=objDatosBDD.RecordCount

if numero>1 then %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY>

<P align="center"><b><font color="#FF0000">LO SIENTO</font></b></P>
<P align="center"><b><font color="#FF0000">EL NUMERO DE HISTORIA CLINICA INGRESADO YA EXISTE</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="modificacion_cardex.asp" target=principal>VUELVALO A INTENTAR</a></font></P>

<P align="center">&nbsp;</P>

</BODY>
</HTML>

<%
objDatosBDD.Close
else


IF historia_clinica <> "" THEN
historia_clinica=TRIM(historia_clinica)
ELSE
historia_clinica=""
END IF

IF apellido_paterno <> "" THEN
apellido_paterno=TRIM(apellido_paterno)
ELSE
apellido_paterno=""
END IF

IF apellido_materno <> "" THEN
apellido_materno=TRIM(apellido_materno)
ELSE
apellido_materno=""
END IF

IF nombres <> "" THEN
nombres=LTRIM(nombres)
ELSE
nombres=""
END IF

IF nombre_padre <> "" THEN
nombre_padre=LTRIM(nombre_padre)
ELSE
nombre_padre=""
END IF

IF nombre_madre <> "" THEN
nombre_madre=LTRIM(nombre_madre)
ELSE
nombre_madre=""
END IF

IF fecha_nacimiento <> "" THEN
fecha_nacimiento=fecha_nacimiento
ELSE
fecha_nacimiento=""
END IF

IF fecha_ingreso <> "" THEN
fecha_ingreso=fecha_ingreso
ELSE
fecha_ingreso="100/01/01"
END IF

IF estado <> "" THEN
estado=estado
ELSE
estado=""
END IF

IF sexo <> "" THEN
sexo=sexo
ELSE
sexo=""
END IF

IF localidad <> "" THEN
localidad=localidad
ELSE
localidad=""
END IF




objEnlaceBDD.ConsultarSQL "UPDATE CARDEX SET APELLIDO_PATERNO='"&apellido_paterno&"',APELLIDO_MATERNO='"&apellido_materno&"',NOMBRES='"&nombres&"',NOMBRE_PADRE='"&nombre_padre&"',NOMBRE_MADRE='"&nombre_madre&"',FECHA_NACIMIENTO='"&fecha_nacimiento&"',NUM_CEDULA='"&cedula&"',SEXO='"&sexo&"',LOCALIDAD='"&localidad&"',ESTADO='"&estado&"' where NUM_HISTORIA_CLINICA = '"&historia_clinica&"' "


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
SUS DATOS HAN SIDO MODIFICADOS. <a href="modificacion_cardex.asp" target=principal>MODIFICAR OTRO REGISTRO</a></font></P>

</BODY>
</HTML>
<%
end if 
end if 
%>


