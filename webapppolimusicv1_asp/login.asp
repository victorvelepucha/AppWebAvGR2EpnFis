<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim errMsg, uname, pwd, nextUrl, target, rs, isPost

errMsg  = ""
uname   = ""
target  = ""
nextUrl = Trim(CStr(Request.QueryString("next") & ""))
If Len(nextUrl) = 0 Then nextUrl = Trim(CStr(Request.Form("next") & ""))

If IsLogged() Then GoToUrl AppRoot() & "home.asp"

isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")

If isPost Then
    uname = Trim(CStr(Request.Form("username") & ""))
    pwd   = CStr(Request.Form("password") & "")

    If Len(uname) = 0 Or Len(pwd) = 0 Then
        errMsg = "Debe ingresar el usuario y la contrasena."
    ElseIf Len(uname) > 50 Then
        errMsg = "El nombre de usuario es demasiado largo."
    Else
        Set rs = Q("SELECT ID_USER, USERNAME, PASSWORD, USER_TYPE, USER_PHOTO, EMAIL FROM TBL_USER WHERE USERNAME = ?", _
                   Array("s:" & uname))
        If rs.EOF Then
            errMsg = "Usuario o contrasena incorrectos."
            rs.Close
        Else
            If CheckPassword(pwd, rs("PASSWORD")) Then
                StartSession rs
                rs.Close
                If Len(nextUrl) > 0 And Left(nextUrl, 1) = "/" And InStr(nextUrl, "//") = 0 And InStr(nextUrl, ":") = 0 Then
                    target = nextUrl
                Else
                    target = AppRoot() & "home.asp"
                End If
                GoToUrl target
            Else
                errMsg = "Usuario o contrasena incorrectos."
                rs.Close
            End If
        End If
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesion | <%=H(APP_NAME)%></title>
<link rel="stylesheet" href="<%=H(AppRoot())%>assets/css/style.css">
</head>
<body>
<div class="auth-wrap">
  <div class="auth-card">

    <div class="auth-brand">
      <div class="brand-logo">PM</div>
      <div class="brand-name">Poli<span>Music</span></div>
    </div>
    <div class="auth-sub">Inicia sesion para escuchar tu musica</div>

    <% If Len(errMsg) > 0 Then %>
      <div class="alert alert-error"><%=H(errMsg)%></div>
    <% End If %>
    <% If Len(Request.QueryString("err") & "") > 0 Then %>
      <div class="alert alert-error"><%=H(Request.QueryString("err"))%></div>
    <% End If %>
    <% If Len(Request.QueryString("ok") & "") > 0 Then %>
      <div class="alert alert-ok"><%=H(Request.QueryString("ok"))%></div>
    <% End If %>

    <form method="post" action="<%=H(AppRoot())%>login.asp" autocomplete="off">
      <input type="hidden" name="next" value="<%=H(nextUrl)%>">
      <div class="form-row">
        <label for="username">Usuario</label>
        <input class="input" type="text" id="username" name="username" maxlength="50"
               value="<%=H(uname)%>" placeholder="Tu nombre de usuario" required autofocus>
      </div>
      <div class="form-row">
        <label for="password">Contrasena</label>
        <input class="input" type="password" id="password" name="password" maxlength="50"
               placeholder="Tu contrasena" required>
      </div>
      <button type="submit" class="btn btn-primary"><%=Ico("logout", 16)%> Entrar</button>
    </form>

    <div class="auth-foot">
      No tienes cuenta? <a href="<%=H(AppRoot())%>register.asp">Registrate gratis</a>
    </div>

    <div class="demo-box">
      <b>Cuentas de prueba</b><br>
      Administrador: <code>userAdmin</code> / <code>12345678</code><br>
      Usuario normal: <code>userNormal</code> / <code>12345678</code><br>
      Usuario con clave cifrada: <code>username1</code> / <code>12345678</code>
    </div>

  </div>
</div>
</body>
</html>
<%
If Not IsEmpty(gConn) Then
    On Error Resume Next
    gConn.Close
    Set gConn = Nothing
    On Error GoTo 0
End If
%>
