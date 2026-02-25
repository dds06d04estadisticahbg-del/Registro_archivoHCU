<html>
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
end if

session("opcion")=Request.Form("opcion_consulta")
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Modificación Cardex</title>
</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">ELIMINACION CARDEX</font><br>
</b>
<form method="POST" name="INGRESARDATO" action="reseliminacion_cardex.asp">
  
  <div align="center">
    <center>
  <table border="1" width="57%" bgcolor="#9999FF">
    <tr>
      <td width="116%" align="right" bgcolor="#3366FF">
        <p align="right"><font color="#FFFFFF" face="Agency FB"><b>Ingrese el/la
        <%=session("opcion")%> que desea consultar:&nbsp;&nbsp;</b></font> </td>
      <td width="76%"><input type="text" name="dato_consulta" size="40" tabindex="1"></td>
    </tr>
  </table>
    </center>
  </div>
  
  <p align="center"><input type="submit" value="Consultar" name="elegir" tabindex="2"></p>
</form>
<p>&nbsp;</p>
<p><br>
</p>

</body>

</html>
