
<% 
  session("usuario")=Request.Form ("usuario") 
  session("clave")=Request.Form ("clave")
  txtusuario=session("usuario")	
  txtclave=session("clave")
  
	
	public objEnlaceBDD
	public objDatosBDD

	Set objEnlaceBDD = Server.CreateObject("EnlaceBDD.clsEnlaceBDD")
	Set objDatosBDD = Server.CreateObject("ADODB.Recordset")
	Set objDatosBDD1 = Server.CreateObject("ADODB.Recordset")

	if objEnlaceBDD.Conectar("") then
		'Estado de la conexion: desconectado
		Error
	end if

	objEnlaceBDD.Consultar "usuario", "",  "usuario='" & txtusuario & "' AND clave='" & txtclave & "' AND ESTADO='Ac'", ""   
			
    Set objDatosBDD = objEnlaceBDD.Resultado


	objEnlaceBDD.Consultar "INSTITUCION", "",  "PREDETERMINADO='P'", ""   
			
    Set objDatosBDD1 = objEnlaceBDD.Resultado



if objDatosBDD.EOF or (objDatosBDD.Fields(6) <> "A" and objDatosBDD1.EOF) then
  
  session("bandera")=0
  Response.Buffer=True
  response.clear
  response.redirect "usuario.asp"
  response.end
  objDatosBDD1.Close
  else

  if objDatosBDD1.EOF then
  session("bandera")=-1  
  Response.Buffer=True
  response.clear
  response.redirect "menu_i.asp"
  response.end
  else
  session("cod_usuario")=objDatosBDD.Fields(0)&"-"+objDatosBDD1.Fields(0)
  end if
  session("nombre")=objDatosBDD.Fields(3)
  session("permiso")=objDatosBDD.Fields(6)
  session("bandera")=-1
  Response.Buffer=True
  response.clear
  if (session("permiso")<>"A") then
  response.redirect "menu_u.asp"
  else
  response.redirect "menu_a.asp"
  end if
  response.end

  objDatosBDD.Close

end if
%>
