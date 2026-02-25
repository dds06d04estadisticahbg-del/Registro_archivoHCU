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

dato_consultau=Request.Form("dato_consultau")
opcionu=session("opcionu")


if opcionu="nombre" then
temp=TRIM(dato_consultau)
nombre=temp+"%"
usuario="%"
end if


if opcionu="usuario" then
temp=TRIM(dato_consultau)
usuario=temp+"%"
nombre="%"
end if



	objEnlaceBDD.Consultar "usuario", "",  "(NOMBRE like'"& nombre &"') and (USUARIO like '"& usuario &"')", ""   
			
   	Set objDatosBDD = objEnlaceBDD.Resultado

if not objDatosBDD.EOF then

%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Eliminación de Usuario</title>
</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">ELIMINACION USUARIO</font><br>
</b>
<br>
</p>


<table border="1" cellpadding="0" cellspacing="0" style="border-width:3; border-collapse: collapse" bordercolor="#000080" width="100%" id=tabla1 height="65">
    
       <tr>
    <td width="30%" align="center" bgcolor="#0000ff" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Usuario</b></font>&nbsp;</td>
    <td width="70%" align="center" bgcolor="#0000FF" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Nombre</b></font>&nbsp;</td>
          
       </tr>



       
	<%
	objDatosBDD.MoveFirst  
	i=0
	  do while not objDatosBDD.eof
	  i=i+1
	%>
    <tr>
    <td width="30%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b> <a href="foreliminacion_usuario_cardex.asp?codigou=<%=objDatosBDD.Fields(0)%>"><%=objDatosBDD.Fields(1)%></a></b></font>&nbsp;</td>
    <td width="70%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(3)%></b></font>&nbsp;</td>
    </tr>


  <%
  objDatosBDD.MoveNext
  loop
  objDatosBDD.close
  %>

    
</table>	

</body>

</html>

<%

else

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Eliminación de Usuario</title>
</head>

<body>

<p>
<b><font color="#FF0000" face="Abadi MT Condensed Light">ELIMINACION USUARIO</font><br>
</b>
<br><br><br><br>
<center>No se han encontrado registros para su consulta, <a href="eliminacion_usuario_cardex.asp">vuélvalo a intentar</a>  por favor! </center>
</p>


</body>

</html>


<%

end if

%>