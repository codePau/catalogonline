<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Mostrar Reservas</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript">
function comprobar(){
	correcto=true;
	IDLibro= document.elim.IDLibro.value;

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
			&gt;<a href="MostrarReservas.asp">Reservas</a>
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
	If session("Usuarix") <> "Paula" then
		SQL = "Select * From Reservas where Usuarix= '"&session("Usuarix")&"'"
	Else
		SQL = "Select * From Reservas"
	End If

	Set oRS = oConn.Execute(SQL)

	If oRS.EOF then
		If session("Usuarix") <> "Paula" then
			Response.Write("<script type=""text/javascript"">alert('No existe ninguna reserva');window.location.href='ÁreaPrivada.htm';</script>")
		Else
			Response.Write("<script type=""text/javascript"">alert('No existe ninguna reserva');window.location.href='Admin.htm';</script>")
		End If		
	Else%>
		<table>
		    <tr>
		      <th>Fecha de Reserva</th>
			<%If session("Usuarix")="Paula" then%>
		      <th>Usuarix</th>
			<%End If%>
		      <th>Título</th>
		
		      <th>IDLibro</th>
		            	
		    </tr><% Do Until oRS.EOF %>
		
		    <tr>
		      <td><% = oRS("FechaReserva") %></td>
			<%If session("Usuarix")="Paula" then%>
		      <td><% = oRS("Usuarix") %></td>
			<%End If%>
		      <td><% = oRS("Título") %></td>
			      
		      <td><% = oRS("IDLibro") %></td>
			
		   </tr><% oRS.MoveNext
		   Loop %>
	  	</table>
	  	<p>Introduce el identificador del libro cuya reserva quieres eliminar
		<form method="post" action="ElimReserva.asp" name="elim" onsubmit="return comprobar(this)">
			<input type="text" name="IDLibro"/><br />
			<input type= "submit" value="Eliminar"/>
	  	</form></p>
	<%End If

	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing%>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>
</html>
