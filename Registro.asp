<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Registro</title>
</head>

<body>
<% 
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"

   	'obtengo los datos de la base
	SQL = "Select * from Alumnxs where Usuarix = '"&Request.form("nUsr")&"';"
	Set oRS = oConn.Execute(SQL)
 
	If Err.Description <> "" then 
	  	Response.Write "<b>Error:  " & Err.Description & "</b>"%>
	  	<a href="Registro.htm">Volver</a>
  <%Else
	If oRS.EOF then	
  		SQL = "Insert into Alumnxs (Nombre, Apellido, Dirección, Teléfono, Correo, Usuarix, Contraseña) values ('" 
  		SQL = SQL &Request.Form("Nombre")&"','"&Request.Form("Apel")&"','"&Request.Form("Dirección")&"',"
  		SQL = SQL &Request.Form("Teléfono")&", '"&Request.Form("Correo")&"','"&Request.Form("nUsr")&"','"&Request.Form("contraseña")&"');"
  		RecordsAffected=0
  		oConn.Execute SQL, RecordsAffected

  		If RecordsAffected > 0 then
 	 		Response.Write("<script type=""text/javascript"">alert('Ya estás registrad@');window.location.href='default.htm';</script>")
 		Else
          	Response.Write "<b>Se ha producido un error al insertar el registro</b>"
          	Response.Redirect "Registro.html"
  		End If
  	Else
  		Response.Write("<script type=""text/javascript"">alert('Ya existe ese nombre de usuari@');window.location.href='Registro.html';</script>")
   	End If
  End If

  oConn.Close 
  Set oRS = Nothing
  Set oConn = Nothing
%>

</body>
</html>
