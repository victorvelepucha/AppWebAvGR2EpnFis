<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim id, isNew, rs, isPost, fName, fLast, fImg, errMsg, fA, fU, i, selected, nSongs

id     = ToInt(Request.QueryString("id"), 0)
If Len(Request.QueryString("id") & "") = 0 Then id = ToInt(Request.Form("id"), 0)
isNew  = (id <= 0)
isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
errMsg = ""

fName = ""
fLast = ""
fImg  = ""

If Not isNew Then
    Set rs = Q("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array("i:" & id))
    If rs.EOF Then
        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/artists.asp?err=" & Server.URLEncode("El artista indicado no existe.")
    End If
    fName = CStr(rs("ARTIST_NAME") & "")
    fLast = CStr(rs("ARTIST_LASTNAME") & "")
    fImg  = CStr(rs("ARTIST_IMAGE") & "")
    nSongs = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array("i:" & id)), 0)
End If

If isPost Then
    fName = Trim(CStr(Request.Form("artist_name") & ""))
    fLast = Trim(CStr(Request.Form("artist_lastname") & ""))
    fImg  = Trim(CStr(Request.Form("artist_image") & ""))

    If Len(fName) = 0 Or Len(fName) > 100 Then
        errMsg = "El nombre del artista es obligatorio (maximo 100 caracteres)."
    ElseIf Len(fLast) > 100 Then
        errMsg = "El apellido no puede superar los 100 caracteres."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            X "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)", _
              Array("s:" & fName, "s:" & fLast, "s:" & fImg)
            SetFlash "Artista """ & fName & """ creado correctamente.", "ok"
        Else
            X "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", _
              Array("s:" & fName, "s:" & fLast, "s:" & fImg, "i:" & id)
            SetFlash "Artista """ & fName & """ actualizado correctamente.", "ok"
        End If
        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/artists.asp"
    End If
End If

fA = ListResFiles("artistimages")
fU = ListResFiles("userimages")

If isNew Then
    PageTitle = "Nuevo artista"
    PageSub   = "Registra un artista para poder asociarlo a las canciones."
Else
    PageTitle = "Editar artista"
    PageSub   = "Modificando <b>" & H(fName & " " & fLast) & "</b> &middot; " & nSongs & " cancion(es) asociada(s)."
End If
PageId     = "admin_artists"
PageKicker = "Administracion"
%>
<!--#include file="../includes/header.asp"-->

<% If Len(errMsg) > 0 Then %><div class="alert alert-error"><%=H(errMsg)%></div><% End If %>

<form class="form" method="post" action="<%=H(AppRoot())%>admin/artist_form.asp">
  <input type="hidden" name="action" value="save">
  <input type="hidden" name="id" value="<%=id%>">

  <div class="form-grid">
    <div class="form-row">
      <label for="artist_name">Nombre *</label>
      <input class="input" type="text" id="artist_name" name="artist_name" maxlength="100" required value="<%=H(fName)%>">
    </div>
    <div class="form-row">
      <label for="artist_lastname">Apellido</label>
      <input class="input" type="text" id="artist_lastname" name="artist_lastname" maxlength="100" value="<%=H(fLast)%>">
    </div>
  </div>

  <div class="form-row">
    <label for="artist_image">Imagen del artista</label>
    <select class="input" id="artist_image" name="artist_image">
      <option value="">Sin imagen (se generara una portada)</option>
      <%
      If IsArray(fA) And UBound(fA) >= 0 Then
          For i = 0 To UBound(fA)
              selected = ""
              If LCase(Right(CStr(fImg), Len(fA(i)))) = LCase(CStr(fA(i))) Then selected = " selected"
      %>
        <option value="~/artistImages/<%=H(fA(i))%>"<%=selected%>><%=H(fA(i))%> (artistas)</option>
      <%
          Next
      End If
      If IsArray(fU) And UBound(fU) >= 0 Then
          For i = 0 To UBound(fU)
              selected = ""
              If LCase(Right(CStr(fImg), Len(fU(i)))) = LCase(CStr(fU(i))) Then selected = " selected"
      %>
        <option value="~/userImages/<%=H(fU(i))%>"<%=selected%>><%=H(fU(i))%> (usuarios)</option>
      <%
          Next
      End If
      %>
    </select>
    <div class="form-hint">
      Coloca tus imagenes en <code>SongFilesAndImages_resources/ArtistImages</code> y apareceran en esta lista.
    </div>
  </div>

  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%=Ico("save", 16)%> Guardar artista</button>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/artists.asp">Cancelar</a>
  </div>
</form>

<%
If Not isNew Then rs.Close
%>
<!--#include file="../includes/footer.asp"-->
