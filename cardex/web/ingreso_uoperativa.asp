<%
opcion=Request.Form("opcion_institucion")
%>
<html>

<head>
<title>Pagina nueva 2</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset rows="133,29%,*" frameborder=0 framespacing="0" border="0">
  <frame name="superior" scrolling="no" noresize target="intermedio" src="cabingreso_institucion_cardex.asp?opcion=<%=opcion%>">
  <frame name="intermedio" target="inferior" src="aux1.htm">
  <frame name="inferior" src="aux1.htm" target="inferior">
  <noframes>
  <body>

  <p>Esta página usa marcos, pero su explorador no los admite.</p>

  </body>
  </noframes>
</frameset>

</html>
