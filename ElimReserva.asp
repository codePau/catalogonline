<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Eliminar Reserva</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
 
 	'obtengo los datos de la base
	SQL = "Select * From Reservas where IDLibro="&Request.form("IDLibro")&";"
	Set oRS = oConn.Execute(SQL)

	If oRS.EOF then
		Response.Write("<script type=""text/javascript"">alert('No existe ningún libro reservado con ese identificador');window.location.href='MostrarReservas.asp';</script>")
	Else
		'elimino la reserva
		SQL = "Delete * from Reservas where IDLibro="&Request.Form("IDLibro")&";"
		Set oRS = oConn.Execute(SQL)

		SQL = "Update Libros set Estado='Disponible' where IDLibro="&Request.Form("IDLibro")&";"
		Set oRS = oConn.Execute(SQL)

		If session("Usuarix") <> "Paula" then
			Response.Write("<script type=""text/javascript"">alert('La reserva ha sido eliminada');window.location.href='ÁreaPrivada.htm';</script>")
		Else
			Response.Write("<script type=""text/javascript"">alert('La reserva ha sido eliminada');window.location.href='Admin.htm';</script>")
		End If				
	End If
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing
%>
</body>
</html>
