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

'Eliminación de datos en cardex

cod_usuario=Request.Form("cod_usuario")


   objEnlaceBDD.ConsultarSQL "DELETE from USUARIO where COD_USUARIO = '"&cod_usuario&"' "
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY>


<P align="center"><b><font color="#FF0000">EL REGISTRO <%=cod_usuario%> HA SIDO ELIMINADO</font></b></P>

<P align="center">&nbsp;</P>

<P align="center"><font size="2" color="#0000FF">
<a href="eliminacion_usuario_cardex.asp" target=principal>VOLVER A ELIMINAR</a></font></P>

<P align="center">&nbsp;</P>

</BODY>
</HTML>

<%
end if 
%>
