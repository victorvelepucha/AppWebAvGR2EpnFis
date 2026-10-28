<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim rsU, email, birth, photo, errMsg, okMsg, action, isPost
Dim pwdCur, pwdNew, pwdNew2, files, i, selected, myCount, myPlays
Dim rsPwd, stored, dup

action = LCase(Trim(CStr(Request.Form("action") & "")))
isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
errMsg = ""
okMsg  = ""

Set rsU = Q("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", _
            Array("i:" & CurrentUserID()))
If rsU.EOF Then
    rsU.Close
    KillSession
    GoToUrl AppRoot() & "login.asp"
End If

email = CStr(rsU("EMAIL") & "")
birth = FmtDateInput(rsU("BIRTHDAY"))
photo = CStr(rsU("USER_PHOTO") & "")

If isPost Then

    If action = "password" Then
        pwdCur  = CStr(Request.Form("pwd_cur") & "")
        pwdNew  = CStr(Request.Form("pwd_new") & "")
        pwdNew2 = CStr(Request.Form("pwd_new2") & "")

        Set rsPwd = Q("SELECT PASSWORD FROM TBL_USER WHERE ID_USER = ?", Array("i:" & CurrentUserID()))
        stored = CStr(rsPwd("PASSWORD") & "")
        rsPwd.Close

        If Not CheckPassword(pwdCur, stored) Then
            errMsg = "La contrasena actual no es correcta."
        ElseIf Len(pwdNew) < 6 Or Len(pwdNew) > 50 Then
            errMsg = "La nueva contrasena debe tener entre 6 y 50 caracteres."
        ElseIf pwdNew <> pwdNew2 Then
            errMsg = "La confirmacion no coincide con la nueva contrasena."
        Else
            X "UPDATE TBL_USER SET PASSWORD = ? WHERE ID_USER = ?", _
              Array("s:" & HashPassword(pwdNew), "i:" & CurrentUserID())
            SetFlash "Tu contrasena se actualizo correctamente.", "ok"
            rsU.Close
            GoToUrl AppRoot() & "profile.asp"
        End If

    Else
        email = Trim(CStr(Request.Form("email") & ""))
        birth = Trim(CStr(Request.Form("birthday") & ""))
        photo = Trim(CStr(Request.Form("photo") & ""))

        If Len(email) = 0 Or Len(email) > 50 Then
            errMsg = "Debe indicar un correo electronico valido."
        ElseIf InStr(email, "@") = 0 Or InStr(email, ".") = 0 Then
            errMsg = "El formato del correo electronico no es valido."
        Else
            dup = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", _
                           Array("s:" & email, "i:" & CurrentUserID())), 0)
            If dup > 0 Then
                errMsg = "Ese correo electronico ya esta en uso por otra cuenta."
            End If
        End If

        If Len(errMsg) = 0 Then
            If Len(birth) = 0 Then
                X "UPDATE TBL_USER SET EMAIL = ?, BIRTHDAY = NULL, USER_PHOTO = ? WHERE ID_USER = ?", _
                  Array("s:" & email, "s:" & photo, "i:" & CurrentUserID())
            Else
                X "UPDATE TBL_USER SET EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ? WHERE ID_USER = ?", _
                  Array("s:" & email, "d:" & birth, "s:" & photo, "i:" & CurrentUserID())
            End If
            Session("UserPhoto") = photo
            Session("Email") = email
            SetFlash "Tu perfil se actualizo correctamente.", "ok"
            rsU.Close
            GoToUrl AppRoot() & "profile.asp"
        End If
    End If
End If

files = ListResFiles("userimages")
myCount = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array("i:" & CurrentUserID())), 0)
myPlays = ToInt(XS("SELECT ISNULL(SUM(S.PLAYS),0) FROM TBL_SONG S", Empty), 0)

PageTitle  = "Mi perfil"
PageId     = "profile"
PageKicker = "Cuenta"
PageSub    = "Administra tus datos personales y tu contrasena."
%>
<!--#include file="includes/header.asp"-->

<div class="hero">
  <div class="hero-art round">
    <img src="<%=H(CurrentUserPhoto())%>" alt="Foto de perfil">
  </div>
  <div class="hero-info">
    <div class="kicker">Perfil</div>
    <h1><%=H(rsU("USERNAME"))%></h1>
    <div class="hero-meta">
      <%=H(rsU("EMAIL"))%> &middot;
      <% If ToInt(rsU("USER_TYPE"), 0) = USER_TYPE_ADMIN Then %>
        <span class="tag tag-green">Administrador</span>
      <% Else %>
        <span class="tag tag-gray">Usuario</span>
      <% End If %>
      &middot; <b><%=myCount%></b> playlists
    </div>
    <div class="hero-actions">
      <a class="btn btn-outline" href="<%=H(AppRoot())%>playlists.asp"><%=Ico("list", 16)%> Mis playlists</a>
      <a class="btn btn-danger" href="<%=H(AppRoot())%>logout.asp"><%=Ico("logout", 16)%> Cerrar sesion</a>
    </div>
  </div>
</div>

<% If Len(errMsg) > 0 Then %><div class="alert alert-error"><%=H(errMsg)%></div><% End If %>

<div class="section">
  <div class="section-head"><h2>Datos personales</h2></div>
  <form class="form" method="post" action="<%=H(AppRoot())%>profile.asp">
    <input type="hidden" name="action" value="save">

    <div class="form-row">
      <label>Nombre de usuario</label>
      <input class="input" type="text" value="<%=H(rsU("USERNAME"))%>" disabled>
      <div class="form-hint">El nombre de usuario no se puede cambiar.</div>
    </div>

    <div class="form-row">
      <label for="email">Correo electronico *</label>
      <input class="input" type="email" id="email" name="email" maxlength="50" required value="<%=H(email)%>">
    </div>

    <div class="form-row">
      <label for="birthday">Fecha de nacimiento</label>
      <input class="input" type="date" id="birthday" name="birthday" value="<%=H(birth)%>">
    </div>

    <div class="form-row">
      <label for="photo">Foto de perfil</label>
      <select class="input" id="photo" name="photo">
        <%
        If IsArray(files) Then
            If UBound(files) >= 0 Then
                For i = 0 To UBound(files)
                    selected = ""
                    If Len(photo) > 0 Then
                        If LCase(Right(CStr(photo), Len(files(i)))) = LCase(CStr(files(i))) Then selected = " selected"
                    End If
        %>
          <option value="~/userImages/<%=H(files(i))%>"<%=selected%>><%=H(files(i))%></option>
        <%
                Next
            End If
        End If
        %>
        <option value=""<% If Len(photo) = 0 Then %> selected<% End If %>>Sin foto (se usara la predeterminada)</option>
      </select>
    </div>

    <div class="form-actions">
      <button type="submit" class="btn btn-primary"><%=Ico("save", 16)%> Guardar cambios</button>
    </div>
  </form>
</div>

<div class="section">
  <div class="section-head"><h2>Cambiar contrasena</h2></div>
  <form class="form" method="post" action="<%=H(AppRoot())%>profile.asp" autocomplete="off">
    <input type="hidden" name="action" value="password">
    <div class="form-row">
      <label for="pwd_cur">Contrasena actual *</label>
      <input class="input" type="password" id="pwd_cur" name="pwd_cur" maxlength="50" required>
    </div>
    <div class="form-grid">
      <div class="form-row">
        <label for="pwd_new">Nueva contrasena *</label>
        <input class="input" type="password" id="pwd_new" name="pwd_new" maxlength="50" required>
      </div>
      <div class="form-row">
        <label for="pwd_new2">Repetir nueva contrasena *</label>
        <input class="input" type="password" id="pwd_new2" name="pwd_new2" maxlength="50" required>
      </div>
    </div>
    <div class="form-actions">
      <button type="submit" class="btn btn-outline"><%=Ico("check", 16)%> Actualizar contrasena</button>
    </div>
  </form>
</div>

<%
rsU.Close
%>
<!--#include file="includes/footer.asp"-->
