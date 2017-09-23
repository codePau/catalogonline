<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Limpiar Encuesta</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"

	'elimino los datos de la encuesta
	SQL = "Delete * from Encuesta;"
	Set oRS = oConn.Execute(SQL)
	Response.Write("<script type=""text/javascript"">alert('Los datos de la encuesta han sido puestos a cero');window.location.href='Admin.htm';</script>")

	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing
%>
</body>
</html>
