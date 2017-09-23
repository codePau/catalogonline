<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Modificar Datos</title>
<link href="biblio.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript">
function comprobar(){
	correcto=true;

	nombre= document.Reg.Nombre.value;
	apellidos= document.Reg.Apel.value;
	dirección= document.Reg.Dirección.value;
	teléfono= document.Reg.Teléfono.value;
	correo= document.Reg.Correo.value;
	
	contraseña= document.Reg.contraseña.value;
	contraseña2= document.Reg.contraseña2.value;

	//Comprueba que ningún campo esté vacío
	if(nombre=="" || apellidos=="" || dirección=="" || teléfono=="" || correo=="" || contraseña=="" || contraseña2==""){
		alert("Debes rellenar todos los campos");
		correcto=false;
	}
	//Si todos los campos están completos, se comprueba que el teléfono es un número, que la dirección de correo tiene el formato adecuado
	// y que las contraseñas coinciden
	else{
		if(isNaN(teléfono)){
			alert("El teléfono debe tener formato numérico");
			correcto=false;
		}
		else{
			if(comprobarCorreo(correo) == -1){	
				alert("La dirección de correo electrónico no es válida");
				correcto=false;
			}
			else if(contraseña != contraseña2){
					alert("Comprueba tu contraseña");
					correcto=false;
			}
		}
	}
return correcto;
}

function comprobarCorreo(correo){
correoOK = 0;
	//Comprueba que sólo haya una arroba y que no esté al principio
	var arroba = correo.indexOf("@",0); 
	if (arroba<1 || correo.lastIndexOf("@") != arroba) correoOK = -1;
		
	//Comprueba que el punto que separa el servidor y el dominio esté en la posición correcta
	var punto = correo.lastIndexOf(".");
	if(punto<arroba || punto==correo.length-1) correoOK = -1;

	return correoOK;
}
</script>
</head>

<body>
<div id="ola">
	<div id="cabecera">
		<div id="logo" style="height: 69px">
			<h1>biblioteca</h1>
			<br />
			<br />
			<br />
			&nbsp;&nbsp;&nbsp;&gt;&gt; Catálogo online
		</div>
		<div id ="cerrar">
			<a href="Cerrar.asp">Cerrar Sesión</a>
		</div>
		<br/>
<%
	'creo conexión 
	Set oConn = Server.CreateObject ("ADODB.Connection")

  	'abro conexión con base de datos
  	oConn.Open "Data Source="& Server.MapPath("biblio.mdb") &";Provider=Microsoft.Jet.OLEDB.4.0"%>
  	

		<div id="miga" style=" width: 403px">
			<a href="default.htm">Área Pública</a>
	<%If session("Usuarix")="Paula" then%>		
			&gt;<a href="Admin.htm">Administradora</a>
			&gt;<a href="MostrarDatos.asp">Seleccionar usuari@</a>
	<%Else%>
			&gt;<a href="ÁreaPrivada.htm">Área Privada</a>
	<%End If%>
			&gt;<a href="ModDatos.asp">Modificar datos de usuari@</a>
		</div>
	</div><!-- fin #cabecera -->
	<div id="menu">
		<ul>
			<li><a href="default.htm">ÁREA PÚBLICA</a></li>
			<li><a href="Catálogo.asp">CATÁLOGO</a></li>
			<li><a href="Mapa.html">MAPA WEB</a></li>
		</ul>
	</div><!-- fin #menu -->
</div><!-- fin #ola -->
	<div id="page" style="height: 390px">
<%
	If session("Usuarix")="Paula" then	
		Usuarix=Request.form("Usuarix")
	Else
		Usuarix=session("Usuarix")
	End If
	'obtengo los datos de la base
	If Usuarix <> "" then
		SQL = "Select * From Alumnxs where Usuarix='"&Usuarix&"';"
		Set oRS = oConn.Execute(SQL)
	
		If oRS.EOF then
			Response.Write("<script type=""text/javascript"">alert('No existe ese nombre de usuari@');window.location.href='MostrarDatos.asp'</script>")
		Else%>
			<form method="post" name="Reg" action="ModDatos2.asp" onsubmit="return comprobar(this)">
				<h2>Datos de contacto</h2>
				Nombre:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="text" name="Nombre" value="<% = oRS("Nombre") %>"/><br />
				Apellido(s):&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="Apellido" value="<% = oRS("Apellido") %>"/><br />
				Dirección:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="Dirección" value="<% = oRS("Dirección") %>"/><br />
				Teléfono:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="Teléfono" value="<% = oRS("Teléfono") %>"/><br />
				Dirección de correo electrónico:<input type="text" name="Correo" value="<% = oRS("Correo") %>"/><br /><br/>
			
				<h2>Datos de la cuenta</h2>
				Contraseña:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="Contraseña" value="<% = oRS("Contraseña") %>"/><br />
	
				<input type="hidden" name="Usuarix" value="<% = oRS("Usuarix") %>"/><br />
				<input type= "submit" value="Enviar"/>
				<input  class="button" type="reset" value="Borrar"/>
			</form>
		<%End If
	Else
		Response.Write("<script type=""text/javascript"">alert('Introduce una usuari@');window.location.href='MostrarDatos.asp';</script>")
	End If
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing
%>
</div><!-- fin #page -->
<div id="pie">Página realizada por Paula Mesa Macías
	<form><input type="button" name="imprimir" value="Imprimir" onclick="window.print();"></form>
</div>
</body>

</html>
