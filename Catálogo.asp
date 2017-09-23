<% Response.CacheControl="No-cache" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="content-type" content="text/html; charset=utf-8" />
  <title>Cat&aacute;logo</title>
  <link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
<div id="ola">
	<div id="cabecera">
		<div id="logo_miga">
			<h1>biblioteca</h1>
			<br />
			<br />
			<br />
			&nbsp;&nbsp;&nbsp;&gt;&gt; Catálogo online
		</div>
		<div id="identificación">
			<form method="get" action="Identificación.asp">
				Nombre de usuari@:<input type="text" name="Nombre"/><br />
				Contraseña:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="password" name="Contraseña"/><br /><br />
				<input type= "submit" value="Identificarse"/>
			</form>
		</div>
		<div id="miga">
			<a href="default.htm">Área Pública</a>
			&gt;<a href="Catálogo.asp">Catálogo</a>
		</div>
	</div><!-- fin #cabecera -->
	<div id="menu">
		<ul>
			<li><a href="default.htm">ÁREA PÚBLICA</a></li>
			<li><a href="Catálogo.asp">CATÁLOGO</a></li>
			<li><a href="Mapa.html">MAPA WEB</a></li>
		</ul>
	</div><!-- fin #menu -->
</div><!-- fin #ola -->
<div id="page">
<%
  On Error Resume Next
  Set oConn = Server.CreateObject("ADODB.Connection")
  Set oRS= Server.CreateObject("ADODB.recordset")
  oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"

  SQL = "Select * From Libros"
  oRS.Open SQL, oConn%>
<table>
    <tr>
      <th>T&iacute;tulo</th>

      <th>Autor(a)</th>

      <th>Editorial</th>

      <th>A&ntilde;oPub</th>
      
    </tr><% Do Until oRS.EOF %>

    <tr>
      <td><% = oRS("Título") %></td>

      <td><% = oRS("Autora") %></td>

      <td><% = oRS("Editorial") %></td>
      
      <td><% = oRS("AñoPub") %></td>

   </tr><% oRS.MoveNext
    Loop %>
  </table><%
 oConn.Close
 Set oRS = Nothing
 Set oConn = Nothing%>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>
</html>
