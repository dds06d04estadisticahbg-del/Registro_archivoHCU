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

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Ingreso Cardex</title>
</head>

<body>
<b>
<font color="#FF0000">&nbsp;&nbsp; REGISTROS</font>
<br>

</b>

<form method="POST" action="ejecutaringreso_cardex.asp">
  <div align="center">
    <center>
    <table border="1" width="55%" height="75">
      
	     <tr>
      <td width="41%" align="center" bgcolor="#3366FF" height="17"><b><a href="eliminacion_cardex.asp"><font face="Agency FB" color="#FFFFFF">Eliminación</font></a></b> </td>
      </tr>
      <tr>
      <td width="41%" align="center" bgcolor="#3366FF" height="20"><b><a href="reporte_cardex.asp"><font face="Agency FB" color="#FFFFFF">Reportes</font></a></b> </td>
      </tr>

    </table>
    </center>
  </div>
</form>




</body>

</html>
