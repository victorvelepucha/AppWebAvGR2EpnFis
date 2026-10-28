<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim nUsers, nAdmins, nSongs, nArtists, nAlbums, nPlaylists, nPlays, nLinks
Dim rsTop, rsUsr, rsRecent

nUsers     = ToInt(XS("SELECT COUNT(*) FROM TBL_USER", Empty), 0)
nAdmins    = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = ?", Array("i:" & USER_TYPE_ADMIN)), 0)
nSongs     = ToInt(XS("SELECT COUNT(*) FROM TBL_SONG", Empty), 0)
nArtists   = ToInt(XS("SELECT COUNT(*) FROM TBL_ARTIST", Empty), 0)
nAlbums    = ToInt(XS("SELECT COUNT(*) FROM TBL_ALBUM", Empty), 0)
nPlaylists = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty), 0)
nPlays     = ToInt(XS("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", Empty), 0)
nLinks     = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG", Empty), 0)

Set rsTop = Q(SqlSongs() & "ORDER BY S.PLAYS DESC, S.SONG_NAME ASC", Empty)
Set rsUsr = Q("SELECT TOP 6 U.ID_USER, U.USERNAME, U.EMAIL, U.USER_TYPE, U.USER_PHOTO, " & _
              "(SELECT COUNT(*) FROM TBL_PLAYLIST P WHERE P.ID_USER = U.ID_USER) AS N " & _
              "FROM TBL_USER U ORDER BY U.ID_USER DESC", Empty)
Set rsRecent = Q("SELECT TOP 5 P.ID_PLAYLIST, P.TITLE, P.CREATION_DATE, P.TYPE, U.USERNAME, " & _
                 "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
                 "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
                 "ORDER BY P.CREATION_DATE DESC, P.ID_PLAYLIST DESC", Empty)

PageTitle  = "Panel de administracion"
PageId     = "admin"
PageKicker = "Modo administrador"
PageSub    = "Gestiona usuarios, canciones, artistas, albumes y playlists de <b>" & H(APP_NAME) & "</b>."
%>
<!--#include file="../includes/header.asp"-->

<div class="section">
  <div class="admin-stats">
    <div class="stat"><span class="ico"><%=Ico("users", 22)%></span><div class="n"><%=nUsers%></div><div class="l">Usuarios</div></div>
    <div class="stat"><span class="ico"><%=Ico("music", 22)%></span><div class="n"><%=nSongs%></div><div class="l">Canciones</div></div>
    <div class="stat"><span class="ico"><%=Ico("mic", 22)%></span><div class="n"><%=nArtists%></div><div class="l">Artistas</div></div>
    <div class="stat"><span class="ico"><%=Ico("disc", 22)%></span><div class="n"><%=nAlbums%></div><div class="l">Albumes</div></div>
    <div class="stat"><span class="ico"><%=Ico("list", 22)%></span><div class="n"><%=nPlaylists%></div><div class="l">Playlists</div></div>
    <div class="stat"><span class="ico"><%=Ico("chart", 22)%></span><div class="n"><%=nPlays%></div><div class="l">Reproducciones</div></div>
  </div>
</div>

<div class="section">
  <div class="section-head"><h2>Accesos rapidos</h2></div>
  <div style="display:flex;gap:10px;flex-wrap:wrap">
    <a class="btn btn-primary" href="<%=H(AppRoot())%>admin/user_form.asp"><%=Ico("plus", 16)%> Nuevo usuario</a>
    <a class="btn btn-primary" href="<%=H(AppRoot())%>admin/song_form.asp"><%=Ico("plus", 16)%> Nueva cancion</a>
    <a class="btn btn-primary" href="<%=H(AppRoot())%>admin/artist_form.asp"><%=Ico("plus", 16)%> Nuevo artista</a>
    <a class="btn btn-primary" href="<%=H(AppRoot())%>admin/album_form.asp"><%=Ico("plus", 16)%> Nuevo album</a>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/users.asp"><%=Ico("users", 16)%> Usuarios</a>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/songs.asp"><%=Ico("music", 16)%> Canciones</a>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/artists.asp"><%=Ico("mic", 16)%> Artistas</a>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/albums.asp"><%=Ico("disc", 16)%> Albumes</a>
    <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/playlists.asp"><%=Ico("list", 16)%> Playlists</a>
  </div>
</div>

<div class="section">
  <div class="section-head"><h2>Canciones mas reproducidas</h2></div>
  <% If rsTop.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("music", 40)%></span><h3>Sin canciones</h3></div>
  <% Else %>
    <div class="table-wrap">
      <table class="table">
        <thead><tr><th>#</th><th>Cancion</th><th>Artistas</th><th>Archivo</th><th>Reproducciones</th><th class="actions">Acciones</th></tr></thead>
        <tbody>
        <%
        Do While Not rsTop.EOF
        %>
        <tr>
          <td><%=rsTop("ID_SONG")%></td>
          <td><b><%=H(rsTop("SONG_NAME"))%></b></td>
          <td><%=H(rsTop("ARTISTAS"))%></td>
          <td>
            <% If SongExists(rsTop("SONG_PATH")) Then %>
              <span class="tag tag-green">Disponible</span>
            <% Else %>
              <span class="tag tag-red">Falta</span>
            <% End If %>
          </td>
          <td><%=ToInt(rsTop("PLAYS"), 0)%></td>
          <td class="actions">
            <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>admin/song_form.asp?id=<%=rsTop("ID_SONG")%>"><%=Ico("edit", 12)%></a>
          </td>
        </tr>
        <%
            rsTop.MoveNext
        Loop
        %>
        </tbody>
      </table>
    </div>
  <% End If %>
  <% rsTop.Close %>
</div>

<div class="section">
  <div class="section-head"><h2>Ultimos usuarios</h2></div>
  <div class="table-wrap">
    <table class="table">
      <thead><tr><th>ID</th><th>Usuario</th><th>Correo</th><th>Tipo</th><th>Playlists</th><th class="actions">Acciones</th></tr></thead>
      <tbody>
      <%
      Do While Not rsUsr.EOF
      %>
      <tr>
        <td><%=rsUsr("ID_USER")%></td>
        <td>
          <div style="display:flex;align-items:center;gap:9px">
            <img class="thumb round" src="<%=H(Resolve(rsUsr("USER_PHOTO"), RES_USERS & "/" & IMG_DEFAULT_USER))%>" alt="">
            <b><%=H(rsUsr("USERNAME"))%></b>
          </div>
        </td>
        <td><%=H(rsUsr("EMAIL"))%></td>
        <td>
          <% If ToInt(rsUsr("USER_TYPE"), 0) = USER_TYPE_ADMIN Then %>
            <span class="tag tag-green">Admin</span>
          <% Else %>
            <span class="tag tag-gray">Usuario</span>
          <% End If %>
        </td>
        <td><%=ToInt(rsUsr("N"), 0)%></td>
        <td class="actions">
          <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>admin/user_form.asp?id=<%=rsUsr("ID_USER")%>"><%=Ico("edit", 12)%></a>
        </td>
      </tr>
      <%
          rsUsr.MoveNext
      Loop
      %>
      </tbody>
    </table>
  </div>
  <% rsUsr.Close %>
</div>

<div class="section">
  <div class="section-head"><h2>Playlists recientes</h2></div>
  <% If rsRecent.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("list", 40)%></span><h3>Sin playlists</h3></div>
  <% Else %>
    <div class="table-wrap">
      <table class="table">
        <thead><tr><th>ID</th><th>Titulo</th><th>Propietario</th><th>Tipo</th><th>Canciones</th><th>Creada</th><th class="actions">Acciones</th></tr></thead>
        <tbody>
        <%
        Do While Not rsRecent.EOF
        %>
        <tr>
          <td><%=rsRecent("ID_PLAYLIST")%></td>
          <td><b><%=H(rsRecent("TITLE"))%></b></td>
          <td><%=H(rsRecent("USERNAME"))%></td>
          <td>
            <% If ToInt(rsRecent("TYPE"), 0) = PLAYLIST_PRIVATE Then %>
              <span class="tag tag-gray">Privada</span>
            <% Else %>
              <span class="tag tag-blue">Publica</span>
            <% End If %>
          </td>
          <td><%=ToInt(rsRecent("N"), 0)%></td>
          <td><%=FmtDate(rsRecent("CREATION_DATE"))%></td>
          <td class="actions">
            <a class="row-btn" title="Ver" href="<%=H(AppRoot())%>playlist.asp?id=<%=rsRecent("ID_PLAYLIST")%>"><%=Ico("search", 12)%></a>
            <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>playlist_edit.asp?id=<%=rsRecent("ID_PLAYLIST")%>"><%=Ico("edit", 12)%></a>
          </td>
        </tr>
        <%
            rsRecent.MoveNext
        Loop
        %>
        </tbody>
      </table>
    </div>
  <% End If %>
  <% rsRecent.Close %>
</div>

<!--#include file="../includes/footer.asp"-->
