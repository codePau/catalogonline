<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Cerrar sesión</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
<div id="ola" style="height: 213px">
	<div id="cabecera">
		<div id="logo">
			<h1>biblioteca</h1><br /><br /><br />&nbsp;&nbsp;&nbsp;&gt;&gt;Catálogo online
		</div>
	</div><!--fin #cabecera -->
</div><!-- fin #ola -->
<div id="page">
	<div id="cerrando">
	<%
		Session.Abandon
		Response.AddHeader "refresh","2;url=default.htm"
	%>
	Cerrando Sesión...
	</div>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>
</html>
