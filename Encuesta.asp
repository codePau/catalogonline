<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Encuesta</title>
</head>

<body>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"
  	
  	'obtengo los datos de la base
	SQL = "Select * from Encuesta"
	Set oRS = oConn.Execute(SQL)
	
	'si es el primer voto, se guarda
	If oRS.EOF then
 		nVotos = 1
		Interes = 25*Request.Form("Interés")-25
		Acceso = 25*Request.Form("Acceso")-25
		Global = 25*Request.Form("Global")-25
				
   	 	SQL = "Insert into Encuesta (p1, p2, p3, nVotos) values("&Interes&","&Acceso&","&Global&", 1);"
  		RecordsAffected=0
  		oConn.Execute SQL, RecordsAffected
			
	'si no, se actualiza el valor de los votos y se pondera para poder representarlo gráficamente
	Else
		nVotos = oRS("nVotos") + 1
		Interes = (25*Request.Form("Interés")-25 + oRS("p1"))/nVotos
		Acceso = (25*Request.Form("Acceso")-25 + oRS("p2"))/nVotos
		Global = (25*Request.Form("Global")-25 + oRS("p3"))/nVotos
	
		SQL = "Update Encuesta set p1="&Interes&", p2 ="&Acceso&", p3 = "&Global&", nVotos = "&nVotos&";"
		oConn.Execute(SQL)		
	End If
  	Response.Write("<script type=""text/javascript"">alert('Tus votos han sido guardados. Gracias!');window.location.href='default.htm';</script>")

	 oConn.Close
	 Set oRS = Nothing
	 Set oConn = Nothing
%>
</body>
</html>
