<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Modificar Datos (2)</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
 
	'actualizo datos de la usuarix	
 	SQL = "Update Alumnxs set Apellido='"&Request.Form("Apellido")&"', Dirección='" 
  	SQL = SQL &Request.Form("Dirección")&"', Teléfono="&Request.Form("Teléfono")&", Correo='"
  	SQL = SQL &Request.Form("Correo")&"',Contraseña='"&Request.Form("Contraseña")&"' where Usuarix='"&Request.Form("Usuarix")&"';"
	RecordsAffected=0
	oConn.Execute(SQL), RecordsAffected
  		
	If RecordsAffected > 0 then
		If session("Usuarix") <> "Paula" then
			Response.Write("<script type=""text/javascript"">alert('Se han actualizado los datos de la usuari@');window.location.href='ÁreaPrivada.htm';</script>")
		Else
			Response.Write("<script type=""text/javascript"">alert('Se han actualizado los datos de la usuari@');window.location.href='Admin.htm';</script>")
		End If
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
