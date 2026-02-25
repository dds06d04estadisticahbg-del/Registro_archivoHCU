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

'Eliminación de datos en cardex

historia_clinica=Request.QueryString("codigoe")

   objEnlaceBDD.ConsultarSQL "DELETE from CARDEX where NUM_HISTORIA_CLINICA = '"&historia_clinica&"' "
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
</HEAD>
<BODY>


<P align="center"><b><font color="#FF0000">EL REGISTRO <%=historia_clinica%> HA SIDO ELIMINADO</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="eliminacion_cardex.asp" target=principal>VOLVER A ELIMINAR</a></font></P>

<P align="center">&nbsp;</P>

</BODY>
</HTML>

<%
end if 
%>


