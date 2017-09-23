<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Añadir Libro</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"

   	'Compruebo que el identificador es único
	SQL = "Select * from Libros where IDLibro="&Request.form("IDLibro")&";"
	Set oRS = oConn.Execute(SQL)
 
	If oRS.EOF then	
  		SQL = "Insert into Libros (Título, Autora, Editorial, AñoPub, Estado, IDLibro) values ('" 
  		SQL = SQL &Request.Form("Título")&"','"&Request.Form("Autora")&"','"&Request.Form("Editorial")&"',"
  		SQL = SQL &Request.Form("AñoPub")&", 'Disponible', "&Request.Form("IDLibro")&");"
  		RecordsAffected=0
  		oConn.Execute SQL, RecordsAffected

  		If RecordsAffected <= 0 then
 	 		Response.Write("<script type=""text/javascript"">alert('Se ha producido un error al insertar el registro');window.location.href='AñadirLibro.htm';</script>")
  		Else 
 	 		Response.Write("<script type=""text/javascript"">alert('Libro añadido a la base de datos');window.location.href='Admin.htm';</script>")
  		End If
  	Else
  		Response.Write("<script type=""text/javascript"">alert('Ya existe ese identificador');window.location.href='AñadirLibro.htm';</script>")
   	End If
  oConn.Close 
  Set oRS = Nothing
  Set oConn = Nothing
%>
</body>
</html>
