<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim id, isNew, rs, isPost, fName, fPath, fPlays, errMsg, files, i
Dim rsArt, rsAlb, selArt, selAlb, arrArt, arrAlb, j, found, sql, base
Dim vArtists, vAlbums

id     = ToInt(Request.QueryString("id"), 0)
If Len(Request.QueryString("id") & "") = 0 Then id = ToInt(Request.Form("id"), 0)
isNew  = (id <= 0)
isPost = (UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
errMsg = ""

fName  = ""
fPath  = ""
fPlays = 0

If Not isNew Then
    Set rs = Q("SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array("i:" & id))
    If rs.EOF Then
        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/songs.asp?err=" & Server.URLEncode("La cancion indicada no existe.")
    End If
    fName  = CStr(rs("SONG_NAME") & "")
    fPath  = CStr(rs("SONG_PATH") & "")
    fPlays = ToInt(rs("PLAYS"), 0)
End If

If isPost Then
    fName  = Trim(CStr(Request.Form("song_name") & ""))
    fPath  = Trim(CStr(Request.Form("song_path") & ""))
    fPlays = ToInt(Request.Form("plays"), 0)

    If Len(fName) = 0 Or Len(fName) > 50 Then
        errMsg = "El nombre de la cancion es obligatorio (maximo 50 caracteres)."
    ElseIf Len(fPath) = 0 Then
        errMsg = "Debe indicar el archivo de audio de la cancion."
    ElseIf fPlays < 0 Then
        errMsg = "Las reproducciones no pueden ser negativas."
    End If

    If Len(errMsg) = 0 Then
        If isNew Then
            X "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?)", _
              Array("s:" & fName, "s:" & fPath, "i:" & fPlays)
            id = ToInt(XS("SELECT MAX(ID_SONG) FROM TBL_SONG", Empty), 0)
            SetFlash "Cancion """ & fName & """ creada correctamente.", "ok"
        Else
            X "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", _
              Array("s:" & fName, "s:" & fPath, "i:" & fPlays, "i:" & id)
            SetFlash "Cancion """ & fName & """ actualizada correctamente.", "ok"
        End If

        ' --- Artistas asociados (TBL_PLAYER) ---
        X "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array("i:" & id)
        vArtists = CStr(Request.Form("artists") & "")
        If Len(vArtists) > 0 Then
            arrArt = Split(vArtists, ",")
            For j = 0 To UBound(arrArt)
                If Len(Trim(arrArt(j))) > 0 Then
                    X "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", _
                      Array("i:" & id, "i:" & ToInt(arrArt(j), 0))
                End If
            Next
        End If

        ' --- Albumes que incluyen la cancion (TBL_TRACKLIST) ---
        X "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array("i:" & id)
        vAlbums = CStr(Request.Form("albums") & "")
        If Len(vAlbums) > 0 Then
            arrAlb = Split(vAlbums, ",")
            For j = 0 To UBound(arrAlb)
                If Len(Trim(arrAlb(j))) > 0 Then
                    X "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", _
                      Array("i:" & ToInt(arrAlb(j), 0), "i:" & id)
                End If
            Next
        End If

        If Not isNew Then rs.Close
        GoToUrl AppRoot() & "admin/songs.asp"
    End If
End If

files = ListResFiles("songfiles")
Set rsArt = Q("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME ASC", Empty)
Set rsAlb = Q("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR FROM TBL_ALBUM ORDER BY TITLE_ALBUM ASC", Empty)

If isNew Then
    PageTitle = "Nueva cancion"
    PageSub   = "Registra una cancion y asociala a sus artistas y albumes."
Else
    PageTitle = "Editar cancion"
    PageSub   = "Modificando <b>" & H(fName) & "</b>."
End If
PageId     = "admin_songs"
PageKicker = "Administracion"
%>
<!--#include file="../includes/header.asp"-->

<% If Len(errMsg) > 0 Then %><div class="alert alert-error"><%=H(errMsg)%></div><% End If %>

<form class="form" method="post" action="<%=H(AppRoot())%>admin/song_form.asp" style="max-width:760px">
  <input type="hidden" name="action" value="save">
  <input type="hidden" name="id" value="<%=id%>">

  <div class="form-grid">
    <div class="form-row">
      <label for="song_name">Nombre de la cancion *</label>
      <input class="input" type="text" id="song_name" name="song_name" maxlength="50" required value="<%=H(fName)%>">
    </div>
    <div class="form-row">
      <label for="plays">Reproducciones iniciales</label>
      <input class="input" type="number" id="plays" name="plays" min="0" value="<%=fPlays%>">
    </div>
  </div>

  <div class="form-row">
    <label for="song_path">Archivo de audio (ruta relativa a los recursos) *</label>
    <input class="input" type="text" id="song_path" name="song_path" required value="<%=H(fPath)%>"
           placeholder="~/songFiles/bensound-adventure.mp3">
    <div class="form-hint">
      Archivos MP3 encontrados en <code>SongFilesAndImages_resources/SongFiles</code>:
      <% If IsArray(files) And UBound(files) >= 0 Then %>
        <span style="display:inline-flex;gap:6px;flex-wrap:wrap;margin-top:8px">
        <%
        For i = 0 To UBound(files)
        %>
          <a class="btn btn-outline btn-sm" href="#" onclick="document.getElementById('song_path').value='~/songFiles/<%=Js(files(i))%>';return false;"><%=H(files(i))%></a>
        <%
        Next
        %>
        </span>
      <% Else %>
        <br>No se encontraron archivos MP3 en la carpeta de canciones.
      <% End If %>
    </div>
  </div>

  <div class="form-row">
    <label for="artists">Artistas interpretes</label>
    <select class="input" id="artists" name="artists" multiple size="6">
      <%
      Do While Not rsArt.EOF
          found = False
          If Not isNew Then
              found = (ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_SONG = ? AND ID_ARTIST = ?", _
                                Array("i:" & id, "i:" & rsArt("ID_ARTIST"))), 0) > 0)
          End If
      %>
        <option value="<%=rsArt("ID_ARTIST")%>"<% If found Then %> selected<% End If %>><%=H(rsArt("ARTIST_NAME") & " " & rsArt("ARTIST_LASTNAME"))%></option>
      <%
          rsArt.MoveNext
      Loop
      %>
    </select>
    <div class="form-hint">Usa Ctrl (o Cmd) para seleccionar varios artistas.</div>
  </div>

  <div class="form-row">
    <label for="albums">Albumes que incluyen la cancion</label>
    <select class="input" id="albums" name="albums" multiple size="6">
      <%
      Do While Not rsAlb.EOF
          found = False
          If Not isNew Then
              found = (ToInt(XS("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_SONG = ? AND ID_ALBUM = ?", _
                                Array("i:" & id, "i:" & rsAlb("ID_ALBUM"))), 0) > 0)
          End If
      %>
        <option value="<%=rsAlb("ID_ALBUM")%>"<% If found Then %> selected<% End If %>><%=H(rsAlb("TITLE_ALBUM"))%> (<%=H(rsAlb("RELEASE_YEAR"))%>)</option>
      <%
          rsAlb.MoveNext
      Loop
      %>
    </select>
  </div>

  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%=Ico("save", 16)%> Guardar cancion</button>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/songs.asp">Cancelar</a>
  </div>
</form>

<%
rsArt.Close
rsAlb.Close
If Not isNew Then rs.Close
%>
<!--#include file="../includes/footer.asp"-->
