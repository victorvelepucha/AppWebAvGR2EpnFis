<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Endpoint AJAX: registra una reproduccion sobre una cancion.
Dim id, ok

Response.ContentType = "text/plain"
ok = 0
id = ToInt(Request.QueryString("id"), 0)

If Not IsLogged() Then
    Response.Write "0"
    Response.End
End If

If id > 0 Then
    If X("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array("i:" & id)) > 0 Then
        ok = 1
    End If
End If

Response.Write ok
Response.End
%>
