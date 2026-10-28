<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Cierra la sesion del usuario y vuelve a la pantalla de acceso.
KillSession
Response.Redirect AppRoot() & "login.asp?ok=" & Server.URLEncode("Tu sesion se cerro correctamente.")
Response.End
%>
