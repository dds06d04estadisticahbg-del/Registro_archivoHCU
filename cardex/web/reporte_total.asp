<%

	public objEnlaceBDD
	public objDatosBDD

	Set objEnlaceBDD = Server.CreateObject("EnlaceBDD.clsEnlaceBDD")
	Set objDatosBDD = Server.CreateObject("ADODB.Recordset")
	Set objDatosBDD1 = Server.CreateObject("ADODB.Recordset")
	
	if objEnlaceBDD.Conectar("") then
		'Estado de la conexion: desconectado
		Response.Redirect "noconectado.asp"
	end if

objEnlaceBDD.Consultar "CARDEX","","ESTADO='A'","NUM_HISTORIA_CLINICA, APELLIDO_PATERNO ASC"
			
Set objDatosBDD = objEnlaceBDD.Resultado

if objDatosBDD.EOF then
%>
<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Reportes</title>
</head>

<body>
</BR>
<p align="center"><font size="3" color="#FF0000"><B>LO SIENTO</B></font></p>
<p align="center"><font size="3" color="#FF0000"><B>NO SE HAN ENCONTRADO REGISTROS</B></font></p>

</body>

</html>

<%
else
Response.ContentType = "application/vnd.ms-excel"
%>
<TABLE border=1 cellPadding=1 cellSpacing=1 width="75%">
<TR>
    <TD><b>NUM_HISTORIA_CLINICA</b></TD>
    <TD><b>APELLIDO_PATERNO</b></TD>
    <TD><b>APELLIDO_MATERNO</b></TD>
    <TD><b>NOMBRES</b></TD>
    <TD><b>NOMBRE_PADRE</b></TD>
    <TD><b>NOMBRE_MADRE</b></TD>
    <TD><b>FECHA_NACIMIENTO</b></TD>
    <TD><b>NUM_CEDULA</b></TD>
    <TD><b>FECHA_REGISTRO</b></TD>

</TR> 

<%    
	do while not objDatosBDD.eof
%>
<TR>
<TD><b><%=objDatosBDD.Fields(0)%></b></TD>
<TD><b><%=objDatosBDD.Fields(1)%></b></TD>
<TD><b><%=objDatosBDD.Fields(2)%></b></TD>
<TD><b><%=objDatosBDD.Fields(3)%></b></TD>
<TD><b><%=objDatosBDD.Fields(4)%></b></TD>
<TD><b><%=objDatosBDD.Fields(5)%></b></TD>    
<TD><b><%=objDatosBDD.Fields(6)%></b></TD>
<TD><b><%=objDatosBDD.Fields(7)%></b></TD>
<TD><b><%=objDatosBDD.Fields(9)%></b></TD>
</TR> 
<%
	objDatosBDD.MoveNext
	loop
	objDatosBDD.Close

%>
    

</TABLE>
<%
end if
%>    


