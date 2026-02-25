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


	objEnlaceBDD.Consultar "usuario", "",  "COD_USUARIO="& codigou &" ", ""   
			
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
	
	var str = document.MODIFICAR.usuario_cardex.value;
	var str1= "ingresado el Usuario";
	document.forms[0].usuario_cardex.focus();
	break;
	
	
	case 2:
	
	var str = document.MODIFICAR.clave_cardex.value;
	var str1= "ingresado la Clave";
	document.forms[0].clave_cardex.focus();
	break;
	
	case 3:

	var str = document.MODIFICAR.nombre_usuario.value;
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
<b><font color="#FF0000" face="Abadi MT Condensed Light">MODIFICACION USUARIO</font>
<br>
</b>

<form method="POST" onsubmit="return submitForms()" name=MODIFICAR action="ejecutarmodificacion_usuario_cardex.asp">
  <div align="center">
    <center>
    <table border="1" width="60%">
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Usuario:</font></b></td>
        <td width="83%"><input type="text" name="usuario_cardex" readonly size="20" tabindex="1" value="<%=objDatosBDD.Fields(1)%>"></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Clave:</font></b></td>
        <td width="83%"><input type="password" name="clave_cardex" size="20" tabindex="2" value="<%=objDatosBDD.Fields(2)%>"></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Nombre:</font></b></td>
        <td width="83%"><input type="text" name="nombre_usuario" size="39" tabindex="3" value="<%=objDatosBDD.Fields(3)%>"></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Descripción:</font></b></td>
        <td width="83%"><input type="text" name="descripcion_usuario" size="39" tabindex="4" value="<%=objDatosBDD.Fields(4)%>"></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Permiso:</font></b></td>
        <td width="83%"><select size="1" name="permiso_usuario" tabindex="5">
        <%
	if objDatosBDD.Fields(6) = "U" then
	%>  
	  <option value="U" selected>Usuario</option>
          <option value="A">Administrador</option>
        <%
	else
	%>  
        <option value="A" selected>Administrador</option>
	<option value="U">Usuario</option>
        <%
	end if
	%>  

          </select></td>
      </tr>
      <tr>
        <td width="41%" align="right" bgcolor="#3366FF"><b><font color="#FFFFFF" face="Agency FB">Estado:</font></b></td>
        <td width="83%"><select size="1" name="estado_usuario" tabindex="6">
        <%
	if objDatosBDD.Fields(5) = "Ac" then
	%>  
           <option value="Ac" selected>Activo</option>
            <option value="In">Inactivo</option>
        <%
	else
	%>  
           <option value="In" selected>Inactivo</option>
           <option value="Ac">Activo</option>
        <%
	end if
	%>  



          </select></td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value="Modificar" name="B1" tabindex="9"><input type="reset" value="Restablecer" name="B2" tabindex="10"></p>
</form>




</body>

</html>
