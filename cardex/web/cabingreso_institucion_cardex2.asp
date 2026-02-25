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

COD_NACPROV_ADM2=Request.Form("COD_NACPROV_ADM2")

opcion=Request.QueryString("opcion")
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>INGRESO INSTITUCION</title>
<script>
function submitForms()
{
if (validar())
return true;
else
return false;
}

function validar()
{
s=document.INGRESARDATO.COD_NACPROV_ADM.value;
if(s==0)
{
alert("Usted no ha escogido una institución. Hágalo por favor!");
document.INGRESARDATO.COD_NACPROV_ADM.focus;
return false;
}
else
return true;
}
</script>

</head>

<body>
      
      
      <form method="POST" name="INGRESARDATO" onsubmit="return submitForms()" action="ejecutaringreso_institucion.asp" target="principal">

  
  <div align="center">
    <center>

  <table border="1" width="77%" bgcolor="#9999FF">
    <tr>



   <td width="104%" align="right" bgcolor="#3366FF">
        <p align="right"><font color="#FFFFFF" face="Agency FB"><b>Ingrese la U. Operativa que desea ingresar:&nbsp;&nbsp;</b></font> </td>     
	<td width="102%">

    <select size="1" name="COD_NACPROV_ADM" tabindex="12">
	<OPTION selected value=0>Seleccione una Institución</OPTION>
	<%
	objEnlaceBDD.Consultar "INSTITUCION", "", "REFERENCIA_INSTITUCION='L' and TIPO_INSTITUCION='OPR' AND COD_LOCAL_ADM='"&COD_NACPROV_ADM2&"' ", ""   
	Set objDatosBDD2 = objEnlaceBDD.Resultado
	
	objDatosBDD2.MoveFirst
	do while not objDatosBDD2.eof
	%>
		<OPTION value='<%=objDatosBDD2.Fields(0)%>'><%=objDatosBDD2.Fields(2)%></OPTION>
	<%
	objDatosBDD2.MoveNext
	Loop
	objDatosBDD2.close
        %>	

	</select>&nbsp; <input type="submit" value="Ingresar" name="elegir" tabindex="4">


</td>


    </tr>
  </table>

    </center>
  </div>
  
  <p align="center">&nbsp;</p>
</form>
<p>&nbsp;</p>
<p><br>
</p>

</body>

</html>
