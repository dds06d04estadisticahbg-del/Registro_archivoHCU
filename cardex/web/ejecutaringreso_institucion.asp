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



'Ingreso de datos en institución

COD_NACPROV_ADM=Request.Form("COD_NACPROV_ADM")

objEnlaceBDD.Consultar "INSTITUCION", "", "COD_NACPROV_ADM='"&COD_NACPROV_ADM&"' and PREDETERMINADO = 'P'", ""   
  
Set objDatosBDD = objEnlaceBDD.Resultado

if not objDatosBDD.eof then %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY>

<P align="center"><b><font color="#FF0000">LO SIENTO</font></b></P>
<P align="center"><b><font color="#FF0000">YA EXISTE UNA INSTITUCION</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="admin_cardex.asp" target=principal>INTENTE EJECUTAR OTRA OPCION</a></font></P>

<P align="center">&nbsp;</P>

</BODY>
</HTML>

<%
objDatosBDD.Close
else


IF COD_NACPROV_ADM <> "" THEN
COD_NACPROV_ADM=COD_NACPROV_ADM
ELSE
COD_NACPROV_ADM=""
END IF

objEnlaceBDD.ConsultarSQL "UPDATE institucion SET PREDETERMINADO = 'P' WHERE COD_NACPROV_ADM='"&COD_NACPROV_ADM&"'"

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
SUS DATOS HAN SIDO REGISTRADOS. <a href="usuario.asp" target=_top>SALGA DEL SISTEMA ANTES DE COMENZAR A TRABAJAR</a></font></P>

</BODY>
</HTML>
<%
end if 
end if 
%>





