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

dato_consulta=Request.Form("dato_consulta")
dato_consulta1=Request.Form("dato_consulta1")
dato_consulta2=Request.Form("dato_consulta2")
opcion=session("opcion")


if opcion="nombre" then
temp=TRIM(dato_consulta)
temp1=TRIM(dato_consulta1)
temp2=TRIM(dato_consulta2)
nombre=temp+"%"
if temp1<>"" then
nombre1=temp1+"%"
else
nombre1="%"
end if
if temp2<>"" then
nombre2=temp2+"%"
else
nombre2="%"
end if
cedula="%"
historia_clinica="%"
fecha_nacimiento="%"
end if


if opcion="cedula" then
temp=TRIM(dato_consulta)
cedula=temp+"%"
nombre="%"
nombre1="%"
nombre2="%"
historia_clinica="%"
fecha_nacimiento="%"
end if


if opcion="historia_clinica" then
temp=TRIM(dato_consulta)
historia_clinica=temp+"%"
cedula="%"
nombre="%"
nombre1="%"
nombre2="%"
fecha_nacimiento="%"
end if

if opcion="fecha_nacimiento" then
temp=TRIM(dato_consulta)
fecha_nacimiento=temp+"%"
cedula="%"
nombre="%"
nombre1="%"
nombre2="%"
historia_clinica="%"
end if





	objEnlaceBDD.Consultar "cardex", "",  "(APELLIDO_PATERNO like '"& nombre &"') and (APELLIDO_MATERNO like '"& nombre1 &"') and (NOMBRES like '"& nombre2 &"') and (NUM_CEDULA like '"& cedula &"') and (NUM_HISTORIA_CLINICA like '"& historia_clinica &"') and (FECHA_NACIMIENTO like '"& fecha_nacimiento &"')", "NUM_HISTORIA_CLINICA, APELLIDO_PATERNO ASC"   
			
   	Set objDatosBDD = objEnlaceBDD.Resultado

if not objDatosBDD.EOF then

numero=objDatosBDD.RecordCount

%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Consulta de Información</title>
</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">CONSULTA CARDEX</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font color="#3366FF" face="Agency FB" size="3">Se han encontrado <%=numero%> registros</font><br>
</b>
<br>
</p>


<table border="1" cellpadding="0" cellspacing="0" style="border-width:3; border-collapse: collapse" bordercolor="#000080" width="100%" id=tabla1 height="65">
    
       <tr>
    <td width="10%" align="center" bgcolor="#0000ff" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Número de Historia Clínica</b></font>&nbsp;</td>
    <td width="30%" align="center" bgcolor="#0000FF" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Nombre</b></font>&nbsp;</td>
    <td width="15%" align="center" bgcolor="#0000ff" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Nombre del Padre</b></font>&nbsp;</td>
    <td width="15%" align="center" bgcolor="#0000FF" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Nombre de la Madre</b></font>&nbsp;</td>
    <td width="10%" align="center" bgcolor="#0000ff" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Fecha de Nacimiento</b></font>&nbsp;</td>
    <td width="10%" align="center" bgcolor="#0000FF" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Cédula</b></font>&nbsp;</td>
    <td width="10%" align="center" bgcolor="#0000ff" bordercolor="#FFFFFF"><font color="#ffffff" size="2"><b>Fecha de Registro</b></font>&nbsp;</td>
    
         
       </tr>



       
	<%
	objDatosBDD.MoveFirst  
	i=0
	  do while not objDatosBDD.eof
	  i=i+1
	%>
    <tr>
    <td width="10%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(0)%></b></font>&nbsp;</td>
    <td width="30%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(1)%>&nbsp;<%=objDatosBDD.Fields(2)%>&nbsp;<%=objDatosBDD.Fields(3)%></b></font>&nbsp;</td>
    <td width="15%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(4)%></b></font>&nbsp;</td>
    <td width="15%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(5)%></b></font>&nbsp;</td>
    <td width="10%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(6)%></b></font>&nbsp;</td>
    <td width="10%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(7)%></b></font>&nbsp;</td>
    <td width="10%" align="center" bgcolor="#CCCCFF" bordercolor="#FFFFFF"><font color="#000000" size="2"><b><%=objDatosBDD.Fields(9)%></b></font>&nbsp;</td>
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
<title>Consulta de Información</title>
</head>

<body>

<p>
<b><font color="#FF0000" face="Abadi MT Condensed Light">CONSULTA CARDEX</font><br>
</b>
<br><br><br><br>
<center>No se han encontrado registros para su consulta, <a href="consulta_cardex.asp">vuélvalo a intentar</a>  por favor! </center>
</p>


</body>

</html>


<%

end if

%>

