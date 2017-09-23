<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Imprimir Recibo</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
<div id="ola">
	<div id="cabecera">
		<div id="logo_miga">
			<h1>biblioteca</h1><br /><br /><br />&nbsp;&nbsp;&nbsp;&gt;&gt;Catálogo online
		</div>
		<div id ="cerrar">
			<a href="Cerrar.asp">Cerrar Sesión</a>
		</div>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"%>
  	

		<div id="miga" style=" width: 403px">
			<a href="default.htm">Área Pública</a>
	<%If session("Usuarix")="Paula" then%>		
			&gt;<a href="Admin.htm">Administradora</a>
	<%Else%>
			&gt;<a href="ÁreaPrivada.htm">Área Privada</a>
	<%End If%>
			&gt;<a href="ReservarLibro.asp">Reservar Libro</a>
			&gt;<a href="Recibo.asp">Recibo</a>
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
  	'obtengo los datos de la base
	If session("Usuarix") <> "Paula" then
		SQL = "Select * From Reservas where Usuarix= '"&session("Usuarix")&"'"
	Else
		SQL = "Select * From Reservas"
	End If
	Set oRS = oConn.Execute(SQL)%>
		<table>
		    <tr>
		      <th>Fecha de Reserva</th>

		      <th>Título</th>
		
		      <th>IDLibro</th>
		            	
		    </tr><% Do Until oRS.EOF %>
		
		    <tr>
		      <td><% = oRS("FechaReserva") %></td>

		      <td><% = oRS("Título") %></td>
			      
		      <td><% = oRS("IDLibro") %></td>
			
		   </tr><% oRS.MoveNext
		   Loop %>
	  	</table>
	  	<br/>
	<form style="height: 25px; width: 76px"><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form><br/>
	<%If session("Usuarix") <> "Paula" then%>
		<form style="height: 25px; width: 76px"><input type="button" name="volver" value="Volver" onclick="window.location.href='ÁreaPrivada.htm'"></form>
	<%Else%>
		<form style="height: 25px; width: 76px"><input type="button" name="volver" value="Volver" onclick="window.location.href='Admin.htm'"></form>
	<%End If
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing%>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
</div>
</body>
</html>
