<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Reservar Libro</title>
</head>

<body>
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
		Else
  			'compruebo su estado
  			If oRS("Estado") <> "Disponible" then
				Response.Write("<script type=""text/javascript"">alert('El libro no está disponible');window.location.href='MostrarLibros.asp';</script>")
			Else
				'guardo el título del libro
				titulo= oRS("Título")
				'y cambio su disponibilidad
				SQL = "Update Libros set Estado='No disponible' where IDLibro="&Request.Form("IDLibro")&";"
				Set oRS = oConn.Execute(SQL)
				If Err.Description <> "" then 
	 				Response.Write "<b>Error:  " & Err.Description & "</b>"%>
	 				<a href="default.htm">Volver</a>
	  			<%Else
	 	 			SQL = "Insert into Reservas (FechaReserva, Usuarix, Título, IDLibro) values (date(), '"&session("Usuarix")&"', '"&titulo&"', "&Request.Form("IDLibro")&")"
					Set oRS = oConn.Execute(SQL)
					Response.Write("<script type=""text/javascript"">alert('El libro ha sido reservado');window.location.href='Recibo.asp';</script>")
				End If
			End If
		End If
	End If
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing
%>
</body>
</html>
