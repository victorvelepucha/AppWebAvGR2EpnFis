<%
' ============================================================================
'  PoliMusic - Autenticacion y control de acceso
' ============================================================================

' ---------------------------------------------------------------------------
'  Estado de la sesion
' ---------------------------------------------------------------------------
Function IsLogged()
    IsLogged = (Len(Session("UserID") & "") > 0)
End Function

Function CurrentUserID()
    CurrentUserID = ToInt(Session("UserID"), 0)
End Function

Function CurrentUserName()
    CurrentUserName = CStr(Session("Username") & "")
End Function

Function CurrentUserType()
    CurrentUserType = ToInt(Session("UserType"), USER_TYPE_NORMAL)
End Function

Function IsAdmin()
    IsAdmin = (IsLogged() And CurrentUserType() = USER_TYPE_ADMIN)
End Function

' Devuelve la URL de la foto del usuario conectado.
Function CurrentUserPhoto()
    Dim p
    p = CStr(Session("UserPhoto") & "")
    CurrentUserPhoto = Resolve(p, RES_USERS & "/" & IMG_DEFAULT_USER)
End Function

' ---------------------------------------------------------------------------
'  Guardas de acceso
' ---------------------------------------------------------------------------

' Exige que haya una sesion iniciada.
Sub RequireLogin()
    If Not IsLogged() Then
        Dim nextUrl
        nextUrl = Request.ServerVariables("URL")
        If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
            nextUrl = nextUrl & "?" & Request.ServerVariables("QUERY_STRING")
        End If
        GoToUrl AppRoot() & "login.asp?next=" & Server.URLEncode(nextUrl)
    End If
End Sub

' Exige privilegios de administrador.
Sub RequireAdmin()
    RequireLogin()
    If Not IsAdmin() Then
        GoToUrl AppRoot() & "home.asp?err=" & Server.URLEncode("No tiene permisos para acceder al area de administracion.")
    End If
End Sub

' ---------------------------------------------------------------------------
'  Inicio / cierre de sesion
' ---------------------------------------------------------------------------
Sub StartSession(rs)
    Session("UserID")   = ToInt(rs("ID_USER"), 0)
    Session("Username") = CStr(rs("USERNAME") & "")
    Session("UserType") = ToInt(rs("USER_TYPE"), USER_TYPE_NORMAL)
    Session("UserPhoto")= CStr(rs("USER_PHOTO") & "")
    Session("Email")    = CStr(rs("EMAIL") & "")
    Session("Flash")    = ""
End Sub

Sub RefreshSession()
    Dim rs
    If Not IsLogged() Then Exit Sub
    Set rs = Q("SELECT ID_USER, USERNAME, USER_TYPE, USER_PHOTO, EMAIL FROM TBL_USER WHERE ID_USER = ?", _
               Array("i:" & CurrentUserID()))
    If rs.EOF Then
        KillSession
    Else
        Session("Username") = CStr(rs("USERNAME") & "")
        Session("UserType") = ToInt(rs("USER_TYPE"), USER_TYPE_NORMAL)
        Session("UserPhoto")= CStr(rs("USER_PHOTO") & "")
        Session("Email")    = CStr(rs("EMAIL") & "")
    End If
    rs.Close
End Sub

Sub KillSession()
    Session.Contents.RemoveAll()
    Session.Abandon
End Sub

' ---------------------------------------------------------------------------
'  Mensajes flash
' ---------------------------------------------------------------------------
Sub SetFlash(msg, kind)
    Session("Flash")     = CStr(msg)
    Session("FlashKind") = CStr(kind)
End Sub

' Devuelve el HTML del mensaje flash pendiente (y lo consume).
Function TakeFlash()
    Dim m, k
    TakeFlash = ""
    m = CStr(Session("Flash") & "")
    k = CStr(Session("FlashKind") & "")
    If Len(m) = 0 Then Exit Function
    Session("Flash") = ""
    Session("FlashKind") = ""
    If Len(k) = 0 Then k = "info"
    TakeFlash = "<div class=""alert alert-" & H(k) & """>" & H(m) & "</div>"
End Function

' Mensaje de error recibido por querystring (para redirecciones).
Function QsAlert()
    Dim e, ok
    QsAlert = ""
    e = CStr(Request.QueryString("err") & "")
    ok = CStr(Request.QueryString("ok") & "")
    If Len(e) > 0 Then QsAlert = QsAlert & "<div class=""alert alert-error"">" & H(e) & "</div>"
    If Len(ok) > 0 Then QsAlert = QsAlert & "<div class=""alert alert-ok"">" & H(ok) & "</div>"
End Function
%>
