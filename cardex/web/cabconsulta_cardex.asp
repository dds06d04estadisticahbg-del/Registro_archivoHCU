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
<title>Consulta Cardex</title>
</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">CONSULTA CARDEX</font><br>
</b>
<form method="POST" name="INGRESARDATO" action="resconsulta_cardex.asp">
  
  <div align="center">
    <center>

<%
if session("opcion")<>"nombre" then
%>

  <table border="1" width="60%" bgcolor="#9999FF">
    <tr>
      <td width="116%" align="right" bgcolor="#3366FF">
        <p align="right"><font color="#FFFFFF" face="Agency FB"><b>Ingrese el
        dato que desea consultar:&nbsp;&nbsp;</b></font> </td>
      <td width="76%"><input type="text" name="dato_consulta" size="40" tabindex="1"></td>
    </tr>
  </table>


<%
else
%>


<table border="1" width="88%" bgcolor="#9999FF" height="31">
    <tr>
      <td width="81%" align="right" bgcolor="#3366FF" height="25">
        <p align="right"><b><font color="#FFFFFF" face="Agency FB">Ingrese el
        dato que desea consultar:&nbsp;&nbsp;</font> </b> </td>
        <td width="32%" height="25"><font face="Agency FB">AP<input type="text" name="dato_consulta" size="15" tabindex="1"></font></td>
	<td width="34%" height="25"><font face="Agency FB">AM<input type="text" name="dato_consulta1" size="15" tabindex="2"></font></td>
	<td width="41%" height="25"><font face="Agency FB">NB<input type="text" name="dato_consulta2" size="30" tabindex="3"></font></td>
    </tr>
  </table>

<%
end if
%>

    </center>
  </div>
  
  <p align="center"><input type="submit" value="Consultar" name="elegir" tabindex="4"></p>
</form>
<p>&nbsp;</p>
<p><br>
</p>

</body>

</html>
