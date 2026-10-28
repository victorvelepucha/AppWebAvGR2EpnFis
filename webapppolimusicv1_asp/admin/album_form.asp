<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim id, isNew, rs, isPost, fTitle, fYear, fImg, errMsg, i, selected, found
Dim fAlb, fPly, rsS, vSongs, arr, j, nTracks

id     = ToInt(Request.QueryString("id"), 0)
If Len(Request.QueryString("id") & "") = 0 Then id = ToInt(Request.Form("id"), 0)
isNew  = (id <= 0)
isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
errMsg = ""

fTitle = ""
fYear  = Year(Date())
fImg   = ""

If Not isNew Then
    Set rs = Q("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array("i:" & id))
    If rs.EOF Then
        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/albums.asp?err=" & Server.URLEncode("El album indicado no existe.")
    End If
    fTitle = CStr(rs("TITLE_ALBUM") & "")
    fYear  = ToInt(rs("RELEASE_YEAR"), 0)
    fImg   = CStr(rs("ALBUM_IMAGE_PATH") & "")
    nTracks = ToInt(XS("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array("i:" & id)), 0)
End If

If isPost Then
    fTitle = Trim(CStr(Request.Form("title_album") & ""))
    fYear  = ToInt(Request.Form("release_year"), 0)
    fImg   = Trim(CStr(Request.Form("album_image") & ""))

    If Len(fTitle) = 0 Or Len(fTitle) > 50 Then
        errMsg = "El titulo del album es obligatorio (maximo 50 caracteres)."
    ElseIf fYear < 1900 Or fYear > 2100 Then
        errMsg = "El anio de lanzamiento debe estar entre 1900 y 2100."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            X "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?)", _
              Array("s:" & fTitle, "i:" & fYear, "s:" & fImg)
            id = ToInt(XS("SELECT MAX(ID_ALBUM) FROM TBL_ALBUM", Empty), 0)
            SetFlash "Album """ & fTitle & """ creado correctamente.", "ok"
        Else
            X "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", _
              Array("s:" & fTitle, "i:" & fYear, "s:" & fImg, "i:" & id)
            SetFlash "Album """ & fTitle & """ actualizado correctamente.", "ok"
        End If

        ' --- Tracklist ---
        X "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array("i:" & id)
        vSongs = CStr(Request.Form("songs") & "")
        If Len(vSongs) > 0 Then
            arr = Split(vSongs, ",")
            For j = 0 To UBound(arr)
                If Len(Trim(arr(j))) > 0 Then
                    X "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", _
                      Array("i:" & id, "i:" & ToInt(arr(j), 0))
                End If
            Next
        End If

        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/albums.asp"
    End If
End If

fAlb = ListResFiles("albumimages")
fPly = ListResFiles("playlistimages")
Set rsS = Q(SqlSongs() & "ORDER BY S.SONG_NAME ASC", Empty)

If isNew Then
    PageTitle = "Nuevo album"
    PageSub   = "Crea un album y selecciona las canciones que lo componen."
Else
    PageTitle = "Editar album"
    PageSub   = "Modificando <b>" & H(fTitle) & "</b> &middot; " & nTracks & " pista(s)."
End If
PageId     = "admin_albums"
PageKicker = "Administracion"
%>
<!--#include file="../includes/header.asp"-->

<% If Len(errMsg) > 0 Then %><div class="alert alert-error"><%=H(errMsg)%></div><% End If %>

<form class="form" method="post" action="<%=H(AppRoot())%>admin/album_form.asp" style="max-width:760px">
  <input type="hidden" name="action" value="save">
  <input type="hidden" name="id" value="<%=id%>">

  <div class="form-grid">
    <div class="form-row">
      <label for="title_album">Titulo del album *</label>
      <input class="input" type="text" id="title_album" name="title_album" maxlength="50" required value="<%=H(fTitle)%>">
    </div>
    <div class="form-row">
      <label for="release_year">Anio de lanzamiento</label>
      <input class="input" type="number" id="release_year" name="release_year" min="1900" max="2100" value="<%=fYear%>">
    </div>
  </div>

  <div class="form-row">
    <label for="album_image">Portada del album</label>
    <select class="input" id="album_image" name="album_image">
      <option value="">Sin portada (se generara una automaticamente)</option>
      <%
      If IsArray(fAlb) And UBound(fAlb) >= 0 Then
          For i = 0 To UBound(fAlb)
              selected = ""
              If LCase(Right(CStr(fImg), Len(fAlb(i)))) = LCase(CStr(fAlb(i))) Then selected = " selected"
      %>
        <option value="~/albumImages/<%=H(fAlb(i))%>"<%=selected%>><%=H(fAlb(i))%> (albumes)</option>
      <%
          Next
      End If
      If IsArray(fPly) And UBound(fPly) >= 0 Then
          For i = 0 To UBound(fPly)
              selected = ""
              If LCase(Right(CStr(fImg), Len(fPly(i)))) = LCase(CStr(fPly(i))) Then selected = " selected"
      %>
        <option value="~/playlistImages/<%=H(fPly(i))%>"<%=selected%>><%=H(fPly(i))%> (playlists)</option>
      <%
          Next
      End If
      %>
    </select>
    <div class="form-hint">Coloca tus imagenes en <code>SongFilesAndImages_resources/AlbumImages</code> y apareceran en esta lista.</div>
  </div>

  <div class="form-row">
    <label for="songs">Canciones del album (tracklist)</label>
    <select class="input" id="songs" name="songs" multiple size="10">
      <%
      Do While Not rsS.EOF
          found = False
          If Not isNew Then
              found = (ToInt(XS("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ? AND ID_SONG = ?", _
                                Array("i:" & id, "i:" & rsS("ID_SONG"))), 0) > 0)
          End If
      %>
        <option value="<%=rsS("ID_SONG")%>"<% If found Then %> selected<% End If %>>
          <%=H(rsS("SONG_NAME"))%> &mdash; <%=H(Shorten(rsS("ARTISTAS"), 40))%>
        </option>
      <%
          rsS.MoveNext
      Loop
      %>
    </select>
    <div class="form-hint">Usa Ctrl (o Cmd) para seleccionar varias canciones.</div>
  </div>

  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%=Ico("save", 16)%> Guardar album</button>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/albums.asp">Cancelar</a>
  </div>
</form>

<%
rsS.Close
If Not isNew Then rs.Close
%>
<!--#include file="../includes/footer.asp"-->
