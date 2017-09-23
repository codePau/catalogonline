<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Identificación</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
  	 
   	'obtengo los datos de la base
	SQL = "Select * from Alumnxs where Usuarix = '"&Request.form("Nombre")&"' and Contraseña = '"&Request.form("Contraseña")&"';"
  	Set oRS = oConn.Execute(SQL)

	If oRS.EOF then	
	  	Response.Write("<script type=""text/javascript"">alert('No existe esa cuenta');window.location.href='default.htm';</script>")
	Else
		session("Usuarix")=oRS("Usuarix")
		If oRS("Usuarix") <> "Paula" then
			Response.Redirect "ÁreaPrivada.htm" 
		Else
			Response.Redirect "Admin.htm"
		End If
  	End If%>
</body>

</html>
