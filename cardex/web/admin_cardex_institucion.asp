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
<title>Ingreso Usuario</title>

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
	
	var str = document.INGRESO.usuario_cardex.value;
	var str1= "ingresado el Usuario";
	document.forms[0].usuario_cardex.focus();
	break;
	
	
	case 2:
	
	var str = document.INGRESO.clave_cardex.value;
	var str1= "ingresado la Clave";
	document.forms[0].clave_cardex.focus();
	break;
	
	case 3:

	var str = document.INGRESO.nombre_usuario.value;
	var str1= "ingresado el Nombre";
	document.forms[0].nombre_usuario.focus();
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
      document.INGRESO.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;

	case 'nf':
	patron = "^[0-9]{"+i+"}$";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se aceptan\n "+i+" dígitos!");
      document.INGRESO.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;
	
	case 'm':
	patron="[a-z0-9_]+@{1}[a-z0-9_]+[a-z0-9_\.]+";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se acepta\n una estructura de la forma\n aaa@bbb.ccc!");
      document.INGRESO.elements[n].select();
      document.forms[0].elements[n].focus();
	  }
	break;
	
	case 'a':
	patron="^[a-záéíóúñ]{0,"+i+"}$";
	var oER = new RegExp(patron,"gi");
	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se acepta un máximo\n de "+i+" caracteres alfabéticos!");
      document.INGRESO.elements[n].select();
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
      document.INGRESO.elements[n].select();
      document.forms[0].elements[n].focus();
	
	}
	else
	{
	var oER = new RegExp(patron,"gi");

	if (!oER.test(v) && v!="")
      {
      alert("\nSólo se aceptan dígitos\ncon máximo "+i+" decimales!");
      document.INGRESO.elements[n].select();
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
<b><font color="#FF0000" face="Abadi MT Condensed Light">INGRESO INSTITUCION</font>
<br>

</b>

<form method="POST" name=INGRESO onsubmit="return submitForms()" action="ejecutaringreso_institucion.asp">
  <div align="center">
    <center>
    <table border="1" width="60%">
      
	     <tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Nombre
        de la Institución:</font></b> </td>
        <td width="83%">

<input type="text" name="NOMBRE_INSTITUCION" size="46" tabindex="1" value=""></td>
      </tr>
      <tr>
      <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Siglas:</font></b> </td>
        <td width="83%">

<input name="SIGLAS_INSTITUCION" size="7" onblur=validar(value,name,'a',6) tabindex="2" value="">
<font face="Agency FB" size="2" color="#3366FF"><b>(máximo 6 letras)</b></font></td>
      </tr>

    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value="Ingresar" name="B1" tabindex="3"><input type="reset" value="Restablecer" name="B2" tabindex="4"></p>
</form>




</body>

</html>
