<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim id, isNew, rs, isPost, action, fUser, fMail, fBirth, fType, fPhoto, fPwd, fPwd2
Dim errMsg, files, i, selected, dup, isSelf

id     = ToInt(Request.QueryString("id"), 0)
If Len(Request.QueryString("id") & "") = 0 Then id = ToInt(Request.Form("id"), 0)
isNew  = (id <= 0)
isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
action = LCase(Trim(CStr(Request.Form("action") & "")))
errMsg = ""

fUser  = ""
fMail  = ""
fBirth = ""
fType  = USER_TYPE_NORMAL
fPhoto = "~/userImages/" & IMG_DEFAULT_USER
fPwd   = ""

If Not isNew Then
    Set rs = Q("SELECT ID_USER, USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", _
               Array("i:" & id))
    If rs.EOF Then
        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/users.asp?err=" & Server.URLEncode("El usuario indicado no existe.")
    End If
    fUser  = CStr(rs("USERNAME") & "")
    fMail  = CStr(rs("EMAIL") & "")
    fBirth = FmtDateInput(rs("BIRTHDAY"))
    fType  = ToInt(rs("USER_TYPE"), USER_TYPE_NORMAL)
    fPhoto = CStr(rs("USER_PHOTO") & "")
End If

isSelf = (Not isNew) And (id = CurrentUserID())

If isPost Then
    fUser  = Trim(CStr(Request.Form("username") & ""))
    fMail  = Trim(CStr(Request.Form("email") & ""))
    fBirth = Trim(CStr(Request.Form("birthday") & ""))
    fType  = ToInt(Request.Form("user_type"), USER_TYPE_NORMAL)
    fPhoto = Trim(CStr(Request.Form("photo") & ""))
    fPwd   = CStr(Request.Form("password") & "")
    fPwd2  = CStr(Request.Form("password2") & "")

    If fType <> USER_TYPE_ADMIN Then fType = USER_TYPE_NORMAL
    If Len(fPhoto) = 0 Then fPhoto = "~/userImages/" & IMG_DEFAULT_USER

    If Len(fUser) < 3 Or Len(fUser) > 50 Then
        errMsg = "El usuario debe tener entre 3 y 50 caracteres."
    ElseIf InStr(fUser, " ") > 0 Then
        errMsg = "El usuario no puede contener espacios."
    ElseIf Len(fMail) = 0 Or Len(fMail) > 50 Or InStr(fMail, "@") = 0 Then
        errMsg = "Debe indicar un correo electronico valido."
    ElseIf isNew And Len(fPwd) < 6 Then
        errMsg = "La contrasena debe tener al menos 6 caracteres."
    ElseIf Len(fPwd) > 0 And Len(fPwd) < 6 Then
        errMsg = "La contrasena debe tener al menos 6 caracteres."
    ElseIf Len(fPwd) > 0 And fPwd <> fPwd2 Then
        errMsg = "Las contrasenas no coinciden."
    ElseIf isSelf And fType <> USER_TYPE_ADMIN Then
        errMsg = "No puedes quitarte a ti mismo los privilegios de administrador."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            dup = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ?", Array("s:" & fUser)), 0)
        Else
            dup = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", Array("s:" & fUser, "i:" & id)), 0)
        End If
        If dup > 0 Then errMsg = "Ese nombre de usuario ya esta en uso."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            dup = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ?", Array("s:" & fMail)), 0)
        Else
            dup = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array("s:" & fMail, "i:" & id)), 0)
        End If
        If dup > 0 Then errMsg = "Ese correo electronico ya esta registrado."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            If Len(fBirth) = 0 Then
                X "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES (?, ?, ?, NULL, ?, ?)", _
                  Array("s:" & fUser, "s:" & HashPassword(fPwd), "s:" & fMail, "s:" & fPhoto, "i:" & fType)
            Else
                X "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES (?, ?, ?, ?, ?, ?)", _
                  Array("s:" & fUser, "s:" & HashPassword(fPwd), "s:" & fMail, "d:" & fBirth, "s:" & fPhoto, "i:" & fType)
            End If
            SetFlash "Usuario """ & fUser & """ creado correctamente.", "ok"
        Else
            If Len(fPwd) > 0 Then
                If Len(fBirth) = 0 Then
                    X "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = NULL, USER_PHOTO = ?, USER_TYPE = ?, PASSWORD = ? WHERE ID_USER = ?", _
                      Array("s:" & fUser, "s:" & fMail, "s:" & fPhoto, "i:" & fType, "s:" & HashPassword(fPwd), "i:" & id)
                Else
                    X "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ?, PASSWORD = ? WHERE ID_USER = ?", _
                      Array("s:" & fUser, "s:" & fMail, "d:" & fBirth, "s:" & fPhoto, "i:" & fType, "s:" & HashPassword(fPwd), "i:" & id)
                End If
            Else
                If Len(fBirth) = 0 Then
                    X "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = NULL, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?", _
                      Array("s:" & fUser, "s:" & fMail, "s:" & fPhoto, "i:" & fType, "i:" & id)
                Else
                    X "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?", _
                      Array("s:" & fUser, "s:" & fMail, "d:" & fBirth, "s:" & fPhoto, "i:" & fType, "i:" & id)
                End If
            End If
            If isSelf Then
                Session("Username") = fUser
                Session("UserPhoto") = fPhoto
            End If
            SetFlash "Usuario """ & fUser & """ actualizado correctamente.", "ok"
        End If
        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/users.asp"
    End If
End If

files = ListResFiles("userimages")

If isNew Then
    PageTitle  = "Nuevo usuario"
    PageSub    = "Crea una cuenta y define su tipo de acceso."
Else
    PageTitle  = "Editar usuario"
    PageSub    = "Modificando la cuenta <b>" & H(fUser) & "</b>."
End If
PageId     = "admin_users"
PageKicker = "Administracion"
%>
<!--#include file="../includes/header.asp"-->

<% If Len(errMsg) > 0 Then %><div class="alert alert-error"><%=H(errMsg)%></div><% End If %>

<form class="form" method="post" action="<%=H(AppRoot())%>admin/user_form.asp" autocomplete="off">
  <input type="hidden" name="action" value="save">
  <input type="hidden" name="id" value="<%=id%>">

  <div class="form-grid">
    <div class="form-row">
      <label for="username">Usuario *</label>
      <input class="input" type="text" id="username" name="username" maxlength="50" required value="<%=H(fUser)%>">
    </div>
    <div class="form-row">
      <label for="email">Correo electronico *</label>
      <input class="input" type="email" id="email" name="email" maxlength="50" required value="<%=H(fMail)%>">
    </div>
  </div>

  <div class="form-grid">
    <div class="form-row">
      <label for="birthday">Fecha de nacimiento</label>
      <input class="input" type="date" id="birthday" name="birthday" value="<%=H(fBirth)%>">
    </div>
    <div class="form-row">
      <label for="user_type">Tipo de usuario</label>
      <select class="input" id="user_type" name="user_type"<% If isSelf Then %> disabled<% End If %>>
        <option value="<%=USER_TYPE_NORMAL%>"<% If fType = USER_TYPE_NORMAL Then %> selected<% End If %>>Usuario (0)</option>
        <option value="<%=USER_TYPE_ADMIN%>"<% If fType = USER_TYPE_ADMIN Then %> selected<% End If %>>Administrador (1)</option>
      </select>
      <% If isSelf Then %>
        <input type="hidden" name="user_type" value="<%=USER_TYPE_ADMIN%>">
        <div class="form-hint">No puedes cambiar tu propio tipo de usuario.</div>
      <% End If %>
    </div>
  </div>

  <div class="form-row">
    <label for="photo">Foto de perfil</label>
    <select class="input" id="photo" name="photo">
      <%
      If IsArray(files) Then
          If UBound(files) >= 0 Then
              For i = 0 To UBound(files)
                  selected = ""
                  If Len(fPhoto) > 0 Then
                      If LCase(Right(CStr(fPhoto), Len(files(i)))) = LCase(CStr(files(i))) Then selected = " selected"
                  End If
      %>
        <option value="~/userImages/<%=H(files(i))%>"<%=selected%>><%=H(files(i))%></option>
      <%
              Next
          End If
      End If
      %>
    </select>
  </div>

  <div class="form-grid">
    <div class="form-row">
      <label for="password"><% If isNew Then %>Contrasena *<% Else %>Nueva contrasena<% End If %></label>
      <input class="input" type="password" id="password" name="password" maxlength="50"<% If isNew Then %> required<% End If %>>
      <div class="form-hint"><% If isNew Then %>Minimo 6 caracteres.<% Else %>Dejar en blanco para conservar la actual.<% End If %></div>
    </div>
    <div class="form-row">
      <label for="password2">Repetir contrasena</label>
      <input class="input" type="password" id="password2" name="password2" maxlength="50">
    </div>
  </div>

  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%=Ico("save", 16)%> Guardar usuario</button>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/users.asp">Cancelar</a>
  </div>
</form>

<%
If Not isNew Then rs.Close
%>
<!--#include file="../includes/footer.asp"-->
