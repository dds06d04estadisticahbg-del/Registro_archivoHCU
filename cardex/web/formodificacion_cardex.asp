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

historia_clinica=Request.QueryString("codigo")


	objEnlaceBDD.Consultar "cardex", "",  "NUM_HISTORIA_CLINICA='"& historia_clinica &"'", ""   
			
   	Set objDatosBDD = objEnlaceBDD.Resultado



%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Modificación Cardex</title>

   <SCRIPT language=JavaScript>

<!--

   function submitForms()
   {
   if ( (valid(1)) && (valid(2)) && (valid(3)) )
      if (confirm("\nEsta seguro de que los datos son correctos?\n\n Haga click en ACEPTAR para enviar.\n\n Haga click en CANCELAR PARA ABORTAR."))
         {
         return true;
         }
      else
         {
         alert("\nUsted ha escogido abortar el pedido.");
         return false;      
         }
   else
      return false;
  
   }




  function valid(i)
   {
   
   switch (i)
	{
	
	case 1:
	
	var str = document.MODIFICAR.historia_clinica.value;
	var str1= "ingresado el Número de Historia Clínica";
	document.forms[0].historia_clinica.focus();
	break;
	
	
	case 2:
	
	var str = document.MODIFICAR.apellido_paterno.value;
	var str1= "ingresado el Apellido Paterno";
	document.forms[0].apellido_paterno.focus();
	break;
	
	case 3:

	var str = document.MODIFICAR.nombres.value;
	var str1= "ingresado el Nombre";
	document.forms[0].nombres.focus();
	break;
	}
	
    if (str == "" || str == '0')
      {
      alert("\nNo ha "+str1+".\n\nPor favor ingréselo!");
      return false;
      }
    else
 
      return true;

    }


// Validación de datos ingresados
   
    function validar(v,n,t,i)
    //v=value, n=name, t=tipo, i=dimensión
	{
	var oER, patron;

	switch (t)
	{
	case 'n':
	patron = "^[0-9]{0,"+i+"}$";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v)&& v!="")
      {
      alert("\nSólo se acepta un máximo de\nde "+i+" dígitos!");
      document.MODIFICAR.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;

	case 'nf':
	patron = "^[0-9]{"+i+"}$";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se aceptan\n "+i+" dígitos!");
      document.MODIFICAR.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;
	
	case 'm':
	patron="[a-z0-9_]+@{1}[a-z0-9_]+[a-z0-9_\.]+";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se acepta\n una estructura de la forma\n aaa@bbb.ccc!");
      document.MODIFICAR.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;
	
	case 'a':
	patron="^[a-záéíóúñ]+$";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se aceptan\n caracteres alfabéticos!");
      document.MODIFICAR.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;
	
	case 'd':
	patron="^[0-9]+\.[0-9]{"+i+"}$";
	patron1 = "[\.]{1}";
	
	var oER1 = new RegExp(patron1,"gi");

	if(!oER1.test(v) && v!="")
	{
	  alert("\nSólo se aceptan dígitos\ncon máximo "+i+" decimales y un punto!");
      document.MODIFICAR.elements[n].select();
      document.forms[0].elements[n].focus();
	
	}
	else
	{
	var oER = new RegExp(patron,"gi");

	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se aceptan dígitos\ncon máximo "+i+" decimales!");
      document.MODIFICAR.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	 }  
	break;

	case 'f':
	patron = "^[0-9]{4}/(0?[1-9]|1[0-2])/(0?[1-9]|1[0-9]|2|2[0-9]|3[0-1])$";
	var oER = new RegExp(patron,"gi");

	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se acepta una estructura de fecha \ndel tipo aaaa/mm/dd coherente!");
      document.forms[0].elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;

	//patron = "^[a-zA-Z-_]+$"
	//patron = "^[0-9]+\.{1}[0-9]+$";
	//patron="^http[s]?://[A-Z][\.[A-Z]]+$"
	/*patron="[a-z0-9_]+@{1}[a-z0-9_]+[a-z0-9_\.]+";
	var oER = new RegExp(patron,"gi");

	if (oER.test(e))
      alert("\nCorrecto!");
	else
      alert("\nSolo se aceptan letras!");
	*/
	}
	}


-->   
</SCRIPT>


</head>

<body>
<b><font color="#FF0000" face="Abadi MT Condensed Light">MODIFICACION CARDEX</font>
<br>
</b>

<form method="POST" onsubmit="return submitForms()" name=MODIFICAR action="ejecutarmodificacion_cardex.asp">
  <div align="center">
    <center>
    <table border="1" width="60%">
      
	     <tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">No.
        de Historia Clínica</font></b><b><font color="#FFFFFF" face="Agency FB">:</font> </b> </td>
        <td width="42%">

<input type="text" name="historia_clinica" size="20" tabindex="1" readonly value="<%=historia_clinica%>"></td>
    </center>
        <td width="21%" bgcolor="#3366FF" colspan="2">

<p align="right"><b><font color="#FFFFFF" face="Agency FB">Estado:</font></b></td>
    <center>
        <td width="20%">



<select size="1" name="estado" tabindex="2">

<%
select case objDatosBDD.Fields(12)
	
	case "A"
	
%>
  <option selected>A</option>
  <option>P</option>
  <option>E</option>
<%	
	
	
	case "P"
	
%>
  <option selected>P</option>
  <option>A</option>
  <option>E</option>
<%	
	
	case "E"
	
%>
  <option selected>E</option>
  <option>A</option>
  <option>P</option>
<%	
end select

%>


</select></td>
      </tr>
      <tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Ap</font></b><b><font color="#FFFFFF" face="Agency FB">ellido
        Paterno:</font> </b> </td>
        <td width="83%" colspan="4">

<input type="text" name="apellido_paterno" size="20" tabindex="3" value="<%=objDatosBDD.Fields(1)%>"></td>
      </tr>
<tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Apellido
        Materno:</font></b> </td>
        <td width="83%" colspan="4">

<input type="text" name="apellido_materno" size="20" tabindex="4" value="<%=objDatosBDD.Fields(2)%>"></td>
</tr>
<tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Nombres:</font></b> </td>
        <td width="83%" colspan="4">



<input type="text" name="nombres" size="39" tabindex="5" value="<%=objDatosBDD.Fields(3)%>"></td>
</tr>
<tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Nombre
        del Padre:</font></b> </td>
        <td width="83%" colspan="4">



<input type="text" name="nombre_padre" size="39" tabindex="6" value="<%=objDatosBDD.Fields(4)%>"></td>
</tr>
<tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Nombre
        de la Madre:</font></b> </td>
        <td width="83%" colspan="4">



<input type="text" name="nombre_madre" size="39" tabindex="7" value="<%=objDatosBDD.Fields(5)%>"></td>
</tr>
<tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Fecha
        de Nacimiento:</font></b> </td>
        <td width="42%" colspan="2">



<input type="text" name="fecha_nacimiento" onblur=validar(value,name,'f',0) size="12" tabindex="8" value="<%=objDatosBDD.Fields(6)%>"> <font color="#3366FF">aaaa/mm/dd</font></td>
        <td width="21%" bgcolor="#3366FF">



<p align="right"><b><font color="#FFFFFF" face="Agency FB">Sexo:</font></b></td>
        <td width="20%">



<select size="1" name="SEXO" tabindex="9">
<%
if objDatosBDD.Fields(10)<>"F" then
%>
  <option selected>M</option>
  <option>F</option>
<%
else
%>
  <option selected>F</option>
  <option>M</option>
<%
end if
%>
  
  
</select></td>
</tr>
<tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Cédula:</font></b> </td>
        <td width="83%" colspan="4">



<input type="text" name="cedula" size="13" onblur=validar(value,name,'nf',10) tabindex="10" value="<%=objDatosBDD.Fields(7)%>"></td>
</tr>

         <tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Dirección:</font></b> </td>
        <td width="83%" colspan="4">



<input type="text" name="localidad" size="50" tabindex="11" value="<%=objDatosBDD.Fields(11)%>"></td>
         </tr>

    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value="Modificar" name="B1" tabindex="9"><input type="reset" value="Restablecer" name="B2" tabindex="10"></p>
</form>




</body>

</html>
