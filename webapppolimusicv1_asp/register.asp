<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim errMsg, okMsg, fUser, fMail, fBirth, isPost
Dim pwd1, pwd2, re, exists, newId, rsNew

errMsg = ""
okMsg  = ""
fUser  = ""
fMail  = ""
fBirth = ""

If IsLogged() Then GoToUrl AppRoot() & "home.asp"

isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")

If isPost Then
    fUser  = Trim(CStr(Request.Form("username") & ""))
    fMail  = Trim(CStr(Request.Form("email") & ""))
    fBirth = Trim(CStr(Request.Form("birthday") & ""))
    pwd1   = CStr(Request.Form("password") & "")
    pwd2   = CStr(Request.Form("password2") & "")

    If Len(fUser) < 3 Or Len(fUser) > 50 Then
        errMsg = "El usuario debe tener entre 3 y 50 caracteres."
    ElseIf InStr(fUser, " ") > 0 Then
        errMsg = "El usuario no puede contener espacios."
    ElseIf Len(fMail) = 0 Or Len(fMail) > 50 Then
        errMsg = "Debe indicar un correo electronico valido."
    ElseIf Len(pwd1) < 6 Or Len(pwd1) > 50 Then
        errMsg = "La contrasena debe tener entre 6 y 50 caracteres."
    ElseIf pwd1 <> pwd2 Then
        errMsg = "Las contrasenas no coinciden."
    Else
        Set re = New RegExp
        re.Pattern = "^[A-Za-z0-9._%\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
        re.IgnoreCase = True
        If Not re.Test(fMail) Then
            errMsg = "El formato del correo electronico no es valido."
        End If
        Set re = Nothing
    End If

    If Len(errMsg) = 0 Then
        exists = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ?", Array("s:" & fUser)), 0)
        If exists > 0 Then
            errMsg = "Ese nombre de usuario ya esta en uso."
        Else
            exists = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ?", Array("s:" & fMail)), 0)
            If exists > 0 Then
                errMsg = "Ese correo electronico ya esta registrado."
            End If
        End If
    End If

    If Len(errMsg) = 0 Then
        If Len(fBirth) = 0 Then
            X "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
              "VALUES (?, ?, ?, NULL, ?, ?)", _
              Array("s:" & fUser, "s:" & HashPassword(pwd1), "s:" & fMail, _
                    "s:" & "~/userImages/" & IMG_DEFAULT_USER, "i:" & USER_TYPE_NORMAL)
        Else
            X "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
              "VALUES (?, ?, ?, ?, ?, ?)", _
              Array("s:" & fUser, "s:" & HashPassword(pwd1), "s:" & fMail, _
                    "d:" & fBirth, "s:" & "~/userImages/" & IMG_DEFAULT_USER, "i:" & USER_TYPE_NORMAL)
        End If

        Set rsNew = Q("SELECT ID_USER, USERNAME, USER_TYPE, USER_PHOTO, EMAIL FROM TBL_USER WHERE USERNAME = ?", _
                      Array("s:" & fUser))
        If Not rsNew.EOF Then
            StartSession rsNew
            SetFlash "Bienvenido a PoliMusic, " & fUser & ". Tu cuenta fue creada correctamente.", "ok"
        End If
        rsNew.Close
        GoToUrl AppRoot() & "home.asp"
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Crear cuenta | <%=H(APP_NAME)%></title>
<link rel="stylesheet" href="<%=H(AppRoot())%>assets/css/style.css">
</head>
<body>
<div class="auth-wrap">
  <div class="auth-card" style="max-width:480px">

    <div class="auth-brand">
      <div class="brand-logo">PM</div>
      <div class="brand-name">Poli<span>Music</span></div>
    </div>
    <div class="auth-sub">Crea tu cuenta y arma tus propias playlists</div>

    <% If Len(errMsg) > 0 Then %>
      <div class="alert alert-error"><%=H(errMsg)%></div>
    <% End If %>

    <form method="post" action="<%=H(AppRoot())%>register.asp" autocomplete="off">
      <div class="form-row">
        <label for="username">Usuario *</label>
        <input class="input" type="text" id="username" name="username" maxlength="50"
               value="<%=H(fUser)%>" placeholder="Entre 3 y 50 caracteres" required autofocus>
      </div>
      <div class="form-row">
        <label for="email">Correo electronico *</label>
        <input class="input" type="email" id="email" name="email" maxlength="50"
               value="<%=H(fMail)%>" placeholder="tucorreo@ejemplo.com" required>
      </div>
      <div class="form-row">
        <label for="birthday">Fecha de nacimiento</label>
        <input class="input" type="date" id="birthday" name="birthday" value="<%=H(fBirth)%>">
        <div class="form-hint">Opcional. Se usa para mostrarte novedades.</div>
      </div>
      <div class="form-grid">
        <div class="form-row">
          <label for="password">Contrasena *</label>
          <input class="input" type="password" id="password" name="password" maxlength="50"
                 placeholder="Minimo 6 caracteres" required>
        </div>
        <div class="form-row">
          <label for="password2">Repetir contrasena *</label>
          <input class="input" type="password" id="password2" name="password2" maxlength="50"
                 placeholder="Repite la contrasena" required>
        </div>
      </div>
      <button type="submit" class="btn btn-primary"><%=Ico("check", 16)%> Crear mi cuenta</button>
    </form>

    <div class="auth-foot">
      Ya tienes cuenta? <a href="<%=H(AppRoot())%>login.asp">Inicia sesion</a>
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
