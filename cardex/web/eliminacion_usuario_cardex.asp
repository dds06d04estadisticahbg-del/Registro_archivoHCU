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
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Consulta Cardex</title>
</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">ELIMINACION USUARIO</font><br>
</b>
<form method="POST" name="ELEGIROPCIONES" action="cabeliminacion_usuario_cardex.asp">
  
  <div align="center">
    <center>
  <table border="1" width="19%" bgcolor="#9999FF">
    <tr>
      <td width="74%" align="right" bgcolor="#3366FF">
        <p align="right"><font color="#FFFFFF" face="Agency FB"><b>Nombre:&nbsp;&nbsp;</b></font> </td>
      <td width="26%"><input type="radio" name="opcion_consulta" value="nombre" checked tabindex="1"></td>
    </tr>
    <tr>
      <td width="74%" align="right" bgcolor="#3366FF"><font color="#FFFFFF" face="Agency FB"><b>Usuario:&nbsp;&nbsp;</b></font> </td>
      <td width="26%"><input type="radio" name="opcion_consulta" value="usuario" tabindex="2"></td>
    </tr>
  </table>
    </center>
  </div>
  
  <p align="center"><input type="submit" value="Consultar" name="elegir" tabindex="5"></p>
</form>
<p>&nbsp;</p>
<p><br>
</p>

</body>

</html>
