﻿<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Modificar Libro</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript">
function comprobar(){
	correcto=true;

	Título= document.Reg.Título.value;
	Autora= document.Reg.Autora.value;
	Editorial= document.Reg.Editorial.value;
	AñoPub= document.Reg.AñoPub.value;
	IDLibro= document.Reg.IDLibro.value;
	
	//Comprueba que ningún campo esté vacío
	if(Título=="" || Autora=="" || Editorial=="" || AñoPub=="" || IDLibro==""){
		alert("Debes rellenar todos los campos");
		correcto=false;
	}
	//Si todos los campos están completos, se comprueba que el año de publicación y el identificador son números
	else{
		if(isNaN(AñoPub)){
			alert("El año de publicación debe tener formato numérico");
			correcto=false;
		}
		else if(isNaN(IDLibro)){
				alert("El identificador debe tener formato numérico");
				correcto=false;
		}
	}
return correcto;
}
</script>
</head>
<body>
<div id="ola">
	<div id="cabecera">
		<div id="logo" style="height: 69px">
			<h1>biblioteca</h1>
			<br />
			<br />
			<br />
			&nbsp;&nbsp;&nbsp;&gt;&gt; Catálogo online
		</div>
		<div id ="cerrar">
			<a href="Cerrar.asp">Cerrar Sesión</a>
		</div>
		<br/>
		<div id="miga" style=" width: 330px">
			<a href="default.htm">Área Pública</a>
			&gt;<a href="Admin.htm">Administradora</a>
			&gt;<a href="MostrarLibros.asp">Seleccionar Libro</a>
			&gt;<a href="ModLibro.asp">Modificar Libro</a>
		</div>
	</div><!-- fin #cabecera -->
	<div id="menu">
		<ul>
			<li><a href="default.htm">ÁREA PÚBLICA</a></li>
			<li><a href="Catálogo.asp">CATÁLOGO</a></li>
			<li><a href="Mapa.html">MAPA WEB</a></li>
		</ul>
	</div>
</div><!-- fin #menu -->
	<div id="page" style="height: 390px">
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
  	
  	'obtengo los datos de la base
	SQL = "Select * From Libros where IDLibro="&Request.form("IDLibro")&";"
	Set oRS = oConn.Execute(SQL)

	If Err.Description <> "" then
		Response.Write "<b>Error:" & Err.Description & "</b>"%>
		<a href="default.htm">Volver</a>
	<%Else
		If oRS.EOF then
			Response.Write("<script type=""text/javascript"">alert('No existe ningún libro con ese identificador');window.location.href='MostrarLibros.asp';</script>")
		Else%>
			<form method="post" name="Reg" action="ModLibro2.asp" onsubmit="return comprobar(this)">
				<h2>Modifica los datos del libro</h2>
				Título:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="text" name="Título" value="<% = oRS("Título") %>"/><br />
				Autor(a):&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <input type="text" name="Autora" value="<% = oRS("Autora") %>"/><br />
				Editorial:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="Editorial"  value="<% = oRS("Editorial") %>"/><br />
				Año de Publicación:&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="AñoPub" value="<% = oRS("AñoPub") %>"/><br /><br/>
				<input type="hidden" name="IDLibro" value="<% = oRS("IDLibro") %>"/>
				<input type= "submit" value="Enviar"/>
				<input  class="button" type="reset" value="Borrar"/>
			</form>
		<%End If
	End If
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing%>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>
</html>