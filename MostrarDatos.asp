<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Mostrar Datos</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
<div id="ola" style="height: 213px">
	<div id="cabecera">
		<div id="logo_miga">
			<h1>biblioteca</h1><br /><br /><br />&nbsp;&nbsp;&nbsp;&gt;&gt;Catálogo online
		</div>
		<div id ="cerrar">
			<a href="Cerrar.asp">Cerrar Sesión</a>
		</div>
		<div id="miga" style=" width: 288px">
			<a href="default.htm">Área Pública</a>
			&gt;<a href="Admin.htm">Administradora</a>
			&gt;<a href="MostrarDatos.asp">Seleccionar usuari@</a>
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
  	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
  	
 	'obtengo los datos de la base
  	SQL = "Select * From Alumnxs"
	Set oRS = oConn.Execute(SQL)%>
	<table>
    	<tr>
    	  <th>Nombre</th>
	
    	  <th>Apellido</th>
	
    	  <th>Direcci&oacute;n</th>
	
    	  <th>Tel&eacute;fono</th>
    	  
    	  <th>Correo</th>
		
    	  <th>Usuari@</th>
    	  
    	  <th>Contrase&ntilde;a</th>
	
    	</tr><% Do Until oRS.EOF %>

    	<tr>
    	  <td><% = oRS("Nombre") %></td>

    	  <td><% = oRS("Apellido") %></td>

    	  <td><% = oRS("Dirección") %></td>

    	  <td><% = oRS("Teléfono") %></td>
      
    	  <td><% = oRS("Correo") %></td>

    	  <td><% = oRS("Usuarix") %></td>
    	  
    	  <td><% = oRS("Contraseña") %></td>
   
   		</tr><% oRS.MoveNext
   		Loop %>
	</table><%
  oConn.Close
  Set oRS = Nothing
  Set oConn = Nothing%>
	<br/>  
	<p>Introduce la usuari@ que quieres modificar</p>
  	<form method="post" action="ModDatos.asp" style="height: 45px">
  		<input type="text" name="Usuarix"/>
  		<input type= "submit" value="Modificar datos"/>
  	</form>
  	<p>Introduce la usuari@ que quieres eliminar</p>
  	<form method="post" action="ElimAlumnx.asp" style="height: 45px">
  	  <input type="text" name="Usuarix"/>
		<input type= "submit" value="Eliminar"/>
  	</form>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>
</html>