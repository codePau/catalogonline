<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Mostrar Libros</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript">
function comprobar(){
	correcto=true;
	IDLibro= document.mod.IDLibro.value;

	if(IDLibro=="" || isNaN(IDLibro)){
		alert("Introduce un identificador válido");
		correcto=false;
	}
return correcto;
}
function comprobar1(){
	correcto=true;
	IDLibro= document.elim.IDLibro.value;

	if(IDLibro=="" || isNaN(IDLibro)){
		alert("Introduce un identificador válido");
		correcto=false;
	}
return correcto;
}
function comprobar2(){
	correcto=true;
	IDLibro= document.res.IDLibro.value;

	if(IDLibro=="" || isNaN(IDLibro)){
		alert("Introduce un identificador válido");
		correcto=false;
	}
return correcto;
}
</script>
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
			&gt;<a href="MostrarLibros.asp">Seleccionar Libro</a>
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
<div id="page" style="height: 390px">
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
  	
  	'obtengo los datos de la base
	SQL = "Select * From Libros"
	Set oRS = oConn.Execute(SQL)%>

	<table>
		<tr>
	      <th>T&iacute;tulo</th>
	
	      <th>Autor(a)</th>
		
	      <th>Editorial</th>
		
	      <th>A&ntilde;oPub</th>
    	  
	      <th>Estado</th>
    	  
	      <th>Identificador</th>
		<tr><td></td></tr>
		
	    </tr><% Do Until oRS.EOF %>
	
	    <tr>
	      <td><% = oRS("Título") %></td>

	      <td><% = oRS("Autora") %></td>
		
	      <td><% = oRS("Editorial") %></td>
		      
	      <td><% = oRS("AñoPub") %></td>
	
	      <td><% = oRS("Estado") %></td>
	
	      <td><% = oRS("IDLibro") %></td>
	
	   </tr><% oRS.MoveNext
	   Loop %>
	</table>
	<%If session("Usuarix") = "Paula" then%>
		<p>Introduce el identificador del libro que quieres modificar
		<form method="post" action="ModLibro.asp" name="mod" onsubmit="return comprobar(this)" style="height: 49px">
			<input type="text" name="IDLibro"/>
			<input type= "submit" value="Modificar datos"/>
		</form></p>
		<p>Introduce el identificador del libro que quieres eliminar
		<form method="post" action="ElimLibro.asp" name="elim" onsubmit="return comprobar1(this)" style="height: 44px">
		  <input type="text" name="IDLibro"/>
 		<input type= "submit" value="Eliminar"/>
 		</form></p>
 	<%End If

	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing%>
   
	<p>Introduce el identificador del libro que quieres reservar
  	<form method="post" action="ReservarLibro.asp" name="res" onsubmit="return comprobar2(this)" style="height: 49px">
  		<input type="text" name="IDLibro"/>
  		<input type= "submit" value="Reservar Libro"/>
  	</form></p>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>
</html>
