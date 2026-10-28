<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Pagina de entrada: envia al inicio o al login segun el estado de la sesion.
If IsLogged() Then
    Response.Redirect AppRoot() & "home.asp"
Else
    Response.Redirect AppRoot() & "login.asp"
End If
Response.End
%>
