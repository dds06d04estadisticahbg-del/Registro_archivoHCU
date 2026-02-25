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

Dim fechactual1
fechactual1 = Date

Function CFecha ( expr )

Dim iDia, iMes, iAno, iPos, iPos2

iPos = InStr(expr, "/")
iDia = Mid(expr, 1, iPos-1)
iPos2 = iPos+1
iPos = InStr(iPos2, expr, "/")
iMes = Mid(expr, iPos2, iPos-iPos2)
iAno = Mid(expr, iPos+1)

CFecha = iAno+"/"+iMes+"/"+iDia 

End Function 'CFecha

fechactual=Cfecha(fechactual1)


'Ingreso de datos en cardex

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
sexo=Request.Form("sexo")
localidad=Request.Form("localidad")
estado="A"


if cedula<>"" then

objEnlaceBDD.Consultar "CARDEX", "NUM_HISTORIA_CLINICA", "NUM_HISTORIA_CLINICA='" & historia_clinica & "' or NUM_CEDULA='" & cedula & "'", ""   

else

objEnlaceBDD.Consultar "CARDEX", "NUM_HISTORIA_CLINICA", "NUM_HISTORIA_CLINICA='" & historia_clinica & "' ", ""   

end if


Set objDatosBDD = objEnlaceBDD.Resultado

if not objDatosBDD.EOF then %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY>

<P align="center"><b><font color="#FF0000">LO SIENTO</font></b></P>
<P align="center"><b><font color="#FF0000">EL NUMERO DE HISTORIA CLINICA INGRESADO YA EXISTE</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="ingreso_cardex.asp" target=principal>VUELVALO A INTENTAR</a></font></P>

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
fecha_nacimiento=TRIM(fecha_nacimiento)
ELSE
fecha_nacimiento=""
END IF

IF fecha_ingreso <> "" THEN
fecha_ingreso=fecha_ingreso
ELSE
fecha_ingreso="100/01/01"
END IF

IF localidad <> "" THEN
localidad=TRIM(localidad)
ELSE
localidad=""
END IF


Dim datos_cardex(13)

datos_cardex(0)="'"& historia_clinica &"'"
datos_cardex(1)="'"& apellido_paterno &"'"
datos_cardex(2)="'"& apellido_materno &"'"
datos_cardex(3)="'"& nombres &"'"
datos_cardex(4)="'"& nombre_padre &"'"
datos_cardex(5)="'"& nombre_madre &"'"
datos_cardex(6)="'"& fecha_nacimiento &"'"
datos_cardex(7)="'"& cedula &"'"
datos_cardex(8)="'"& cod_usuario &"'"
datos_cardex(9)="'"& fecha_ingreso &"'"
datos_cardex(10)="'"& sexo &"'"
datos_cardex(11)="'"& localidad &"'"
datos_cardex(12)="'"& estado &"'"

objEnlaceBDD.AlmacenarSql "cardex", datos_cardex, 13


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
SUS DATOS HAN SIDO REGISTRADOS. <a href="ingreso_cardex.asp" target=principal>VOLVER A INGRESAR</a></font></P>

</BODY>
</HTML>
<%
end if 
end if 
%>




