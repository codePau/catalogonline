<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Modificar Libro (2)</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
 
	'actualizo datos del libro 	
 	SQL = "Update Libros set Título='"&Request.Form("Título")&"', Autora='"&Request.Form("Autora")&"', Editorial='" 
  	SQL = SQL &Request.Form("Editorial")&"', AñoPub="&Request.Form("AñoPub")&" where IDLibro="&Request.Form("IDLibro")&";"
	RecordsAffected=0
	oConn.Execute(SQL), RecordsAffected
  		
	If RecordsAffected > 0 then
		Response.Write("<script type=""text/javascript"">alert('Se han actualizado los datos del libro');window.location.href='Admin.htm';</script>")
	Else
		Response.Write "<b>Se ha producido un error al actualizar el registro</b>"%>
 		<a href="default.htm">Volver</a>
  	<%End If

	 oConn.Close
	 Set oRS = Nothing
	 Set oConn = Nothing
%>
</body>
</html>
