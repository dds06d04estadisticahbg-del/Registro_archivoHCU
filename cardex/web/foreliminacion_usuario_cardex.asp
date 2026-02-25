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

codigou=Request.QueryString("codigou")


	objEnlaceBDD.Consultar "usuario", "",  "COD_USUARIO='"& codigou &"'", ""   
			
   	Set objDatosBDD = objEnlaceBDD.Resultado

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Eliminación Cardex</title>

</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">ELIMINAR USUARIO</font>
<br>
</b>

<form method="POST"  name=ELIMINAR action="ejecutar_eliminacion_usuario_cardex.asp">
  <div align="center">
    <center>
    <table border="1" width="60%">
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Usuario:</font></b></td>
        <td width="83%"><%=objDatosBDD.Fields(1)%></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Clave:</font></b></td>
        <td width="83%"><%=objDatosBDD.Fields(2)%></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Nombre:</font></b></td>
        <td width="83%"><%=objDatosBDD.Fields(3)%></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Descripción:</font></b></td>
        <td width="83%"><%=objDatosBDD.Fields(4)%></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Permiso:</font></b></td>
        <td width="83%">
        <%
	if objDatosBDD.Fields(6) = "U" then
	%>  
	Usuario
         <%
	else
	%>  
	Administrador
        <%
	end if
	%>  

          </select></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Estado:</font></b></td>
        <td width="83%">
        <%
	if objDatosBDD.Fields(5) = "Ac" then
	%>  
	Activo
        <%
	else
	%>  
	Inactivo
        <%
	end if
	%>  

          </select></td>
      </tr>
    </table>
    </center>
  </div>
<input type="hidden" name="cod_usuario" size="20" tabindex="1" value="<%=objDatosBDD.Fields(0)%>">

  <p align="center"><input type="submit" value="Eliminar" name="B1" tabindex="9"></p>
</form>




</body>

</html>
