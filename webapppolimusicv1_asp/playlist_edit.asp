<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim id, isNew, rs, title, ptype, photo, errMsg, files, i, selected, action, isPost

id     = ToInt(Request.QueryString("id"), 0)
If Len(Request.QueryString("id") & "") = 0 Then id = ToInt(Request.Form("id"), 0)
isNew  = (id <= 0)
isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
action = LCase(Trim(CStr(Request.Form("action") & "")))

errMsg = ""
title  = ""
ptype  = PLAYLIST_PUBLIC
photo  = ""

' ------------------------------------------------------- Carga de datos actuales
If Not isNew Then
    Set rs = Q("SELECT ID_PLAYLIST, ID_USER, TITLE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array("i:" & id))
    If rs.EOF Then
        rs.Close
        GoToUrl AppRoot() & "playlists.asp?err=" & Server.URLEncode("La playlist solicitada no existe.")
    End If
    If ToInt(rs("ID_USER"), 0) <> CurrentUserID() And Not IsAdmin() Then
        rs.Close
        GoToUrl AppRoot() & "error.asp?code=403&msg=" & Server.URLEncode("Esta playlist pertenece a otro usuario.")
    End If
    title = CStr(rs("TITLE") & "")
    ptype = ToInt(rs("TYPE"), PLAYLIST_PUBLIC)
    photo = CStr(rs("PHOTO") & "")
End If

' ------------------------------------------------------------------- Acciones
If isPost Then

    If action = "delete" Then
        If Not isNew Then
            X "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array("i:" & id)
            X "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array("i:" & id)
            rs.Close
        End If
        SetFlash "La playlist se elimino correctamente.", "ok"
        GoToUrl AppRoot() & "playlists.asp"
    End If

    title = Trim(CStr(Request.Form("title") & ""))
    ptype = ToInt(Request.Form("type"), PLAYLIST_PUBLIC)
    photo = Trim(CStr(Request.Form("photo") & ""))
    If ptype <> PLAYLIST_PRIVATE Then ptype = PLAYLIST_PUBLIC

    If Len(title) = 0 Then
        errMsg = "El titulo de la playlist es obligatorio."
    ElseIf Len(title) > 50 Then
        errMsg = "El titulo no puede superar los 50 caracteres."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            X "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (?, ?, GETDATE(), ?, ?)", _
              Array("i:" & CurrentUserID(), "s:" & title, "i:" & ptype, "s:" & photo)
            id = ToInt(XS("SELECT MAX(ID_PLAYLIST) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array("i:" & CurrentUserID())), 0)
            SetFlash "Playlist """ & title & """ creada correctamente.", "ok"
        Else
            X "UPDATE TBL_PLAYLIST SET TITLE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?", _
              Array("s:" & title, "i:" & ptype, "s:" & photo, "i:" & id)
            rs.Close
            SetFlash "Playlist actualizada correctamente.", "ok"
        End If
        GoToUrl AppRoot() & "playlist.asp?id=" & id
    End If
End If

files = ListResFiles("playlistimages")

If isNew Then
    PageTitle  = "Nueva playlist"
    PageKicker = "Mis playlists"
    PageSub    = "Ponle un nombre, elige quien puede verla y selecciona una portada."
Else
    PageTitle  = "Editar playlist"
    PageKicker = "Mis playlists"
    PageSub    = "Modifica el titulo, la visibilidad o la portada de la playlist <b>" & H(title) & "</b>."
End If
PageId = "playlists"
%>
<!--#include file="includes/header.asp"-->

<% If Len(errMsg) > 0 Then %><div class="alert alert-error"><%=H(errMsg)%></div><% End If %>

<form class="form" method="post" action="<%=H(AppRoot())%>playlist_edit.asp">
  <input type="hidden" name="action" value="save">
  <input type="hidden" name="id" value="<%=id%>">

  <div class="form-row">
    <label for="title">Titulo *</label>
    <input class="input" type="text" id="title" name="title" maxlength="50" required
           value="<%=H(title)%>" placeholder="Mi playlist favorita">
  </div>

  <div class="form-row">
    <label for="type">Visibilidad</label>
    <select class="input" id="type" name="type">
      <option value="<%=PLAYLIST_PUBLIC%>"<% If ptype = PLAYLIST_PUBLIC Then %> selected<% End If %>>Publica - cualquier usuario puede verla</option>
      <option value="<%=PLAYLIST_PRIVATE%>"<% If ptype = PLAYLIST_PRIVATE Then %> selected<% End If %>>Privada - solo yo puedo verla</option>
    </select>
  </div>

  <div class="form-row">
    <label for="photo">Portada</label>
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
        <option value="~/playlistImages/<%=H(files(i))%>"<%=selected%>><%=H(files(i))%></option>
      <%
              Next
          End If
      End If
      %>
      <option value=""<% If Len(photo) = 0 Then %> selected<% End If %>>Sin portada (se generara una automaticamente)</option>
    </select>
    <div class="form-hint">Las imagenes disponibles se leen de <code>SongFilesAndImages_resources/PlaylistImages</code>.</div>
  </div>

  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%=Ico("save", 16)%> Guardar playlist</button>
    <a class="btn btn-outline" href="<% If isNew Then %><%=H(AppRoot())%>playlists.asp<% Else %><%=H(AppRoot())%>playlist.asp?id=<%=id%><% End If %>">Cancelar</a>
  </div>
</form>

<%
If Not isNew Then
    rs.Close
%>
<form method="post" action="<%=H(AppRoot())%>playlist_edit.asp" style="margin-top:34px"
      data-confirm="Se eliminara la playlist y todas sus relaciones con canciones. Deseas continuar?">
  <input type="hidden" name="action" value="delete">
  <input type="hidden" name="id" value="<%=id%>">
  <button type="submit" class="btn btn-danger"><%=Ico("trash", 16)%> Eliminar esta playlist</button>
</form>
<%
End If
%>

<!--#include file="includes/footer.asp"-->
