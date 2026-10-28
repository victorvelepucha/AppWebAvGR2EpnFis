<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim rsTop, rsMy, rsPub, rsArt, rsAlb, pi, cov, url, art
Dim myCount, pubCount, totalPlays, hours

PageTitle  = "Inicio"
PageId     = "home"
PageKicker = "Bienvenido de nuevo"
PageSub    = "Hola <b>" & H(CurrentUserName()) & "</b>, esta es tu biblioteca musical."

Set rsTop = Q(SqlSongs() & "ORDER BY S.PLAYS DESC, S.SONG_NAME ASC", Empty)
Set rsMy  = Q("SELECT TOP 6 P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.CREATION_DATE, " & _
              "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
              "FROM TBL_PLAYLIST P WHERE P.ID_USER = ? ORDER BY P.CREATION_DATE DESC, P.ID_PLAYLIST DESC", _
              Array("i:" & CurrentUserID()))
Set rsPub = Q("SELECT TOP 6 P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, U.USERNAME, " & _
              "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
              "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
              "WHERE P.TYPE = ? ORDER BY P.CREATION_DATE DESC, P.ID_PLAYLIST DESC", _
              Array("i:" & PLAYLIST_PUBLIC))
Set rsArt = Q("SELECT TOP 6 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE, " & _
              "(SELECT COUNT(*) FROM TBL_PLAYER PL WHERE PL.ID_ARTIST = A.ID_ARTIST) AS N " & _
              "FROM TBL_ARTIST A ORDER BY ARTIST_NAME ASC", Empty)
Set rsAlb = Q("SELECT TOP 6 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH, " & _
              "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
              "FROM TBL_ALBUM A ORDER BY RELEASE_YEAR DESC", Empty)

myCount    = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array("i:" & CurrentUserID())), 0)
pubCount   = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE TYPE = ?", Array("i:" & PLAYLIST_PUBLIC)), 0)
totalPlays = ToInt(XS("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", Empty), 0)
%>
<!--#include file="includes/header.asp"-->

<div class="hero">
  <div class="hero-art" style="display:grid;place-items:center;background:linear-gradient(135deg,#1ed760,#0a5c2a)">
    <div style="font-size:74px;line-height:1">&#9835;</div>
  </div>
  <div class="hero-info">
    <div class="kicker">Tu biblioteca</div>
    <h1>Buen momento para<br>escuchar musica</h1>
    <div class="hero-meta">
      <b><%=myCount%></b> playlists propias &nbsp;&middot;&nbsp;
      <b><%=pubCount%></b> playlists publicas &nbsp;&middot;&nbsp;
      <b><%=totalPlays%></b> reproducciones registradas
    </div>
    <div class="hero-actions">
      <button type="button" class="big-play" id="btnPlayAll" title="Reproducir todo"><%=Ico("play", 20)%></button>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>playlist_edit.asp"><%=Ico("plus", 16)%> Nueva playlist</a>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>songs.asp"><%=Ico("music", 16)%> Ver todas las canciones</a>
    </div>
  </div>
</div>

<!-- ===================== Canciones mas escuchadas ===================== -->
<div class="section">
  <div class="section-head">
    <h2>Las mas escuchadas</h2>
    <a class="more" href="<%=H(AppRoot())%>songs.asp">Ver todas</a>
  </div>
  <% If rsTop.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("music", 40)%></span><h3>Sin canciones</h3><p>Aun no hay canciones registradas.</p></div>
  <% Else %>
    <div class="table-wrap">
      <table class="tracklist">
        <thead>
          <tr><th class="idx">#</th><th class="t-cover"></th><th>Cancion</th><th>Reproducciones</th><th class="t-actions"></th></tr>
        </thead>
        <tbody>
        <%
        pi = 0
        Do While Not rsTop.EOF
            url = Resolve(rsTop("SONG_PATH"), "")
            art = CStr(rsTop("ARTISTAS") & "")
            cov = ""
            Response.Write SongRow(pi, rsTop("ID_SONG"), rsTop("SONG_NAME"), art, rsTop("PLAYS"), url, cov, "")
            If Len(url) > 0 Then pi = pi + 1
            rsTop.MoveNext
        Loop
        %>
        </tbody>
      </table>
    </div>
  <% End If %>
  <% rsTop.Close %>
</div>

<!-- ===================== Mis playlists ===================== -->
<div class="section">
  <div class="section-head">
    <h2>Mis playlists</h2>
    <a class="more" href="<%=H(AppRoot())%>playlists.asp">Ver todas</a>
  </div>
  <% If rsMy.EOF Then %>
    <div class="empty">
      <span class="ico"><%=Ico("list", 40)%></span>
      <h3>Todavia no tienes playlists</h3>
      <p>Crea tu primera playlist y organiza tu musica.</p>
      <p style="margin-top:14px"><a class="btn btn-primary" href="<%=H(AppRoot())%>playlist_edit.asp"><%=Ico("plus", 16)%> Crear playlist</a></p>
    </div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsMy.EOF
          cov = Resolve(rsMy("PHOTO"), RES_PLAYLISTS & "/" & IMG_DEFAULT_PLAYLIST)
      %>
      <a class="card" href="<%=H(AppRoot())%>playlist.asp?id=<%=rsMy("ID_PLAYLIST")%>">
        <div class="art">
          <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
          <% Else %><%=CoverPlaceholder(rsMy("TITLE"), "")%><% End If %>
        </div>
        <div class="title"><%=H(rsMy("TITLE"))%></div>
        <div class="meta">
          <%=ToInt(rsMy("N"), 0)%> canciones &middot;
          <% If ToInt(rsMy("TYPE"), 0) = PLAYLIST_PRIVATE Then %>Privada<% Else %>Publica<% End If %>
        </div>
      </a>
      <%
          rsMy.MoveNext
      Loop
      %>
    </div>
  <% End If %>
  <% rsMy.Close %>
</div>

<!-- ===================== Playlists publicas ===================== -->
<div class="section">
  <div class="section-head">
    <h2>Playlists publicas de la comunidad</h2>
    <a class="more" href="<%=H(AppRoot())%>playlists.asp?scope=public">Ver todas</a>
  </div>
  <% If rsPub.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("list", 40)%></span><h3>Sin playlists publicas</h3></div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsPub.EOF
          cov = Resolve(rsPub("PHOTO"), RES_PLAYLISTS & "/" & IMG_DEFAULT_PLAYLIST)
      %>
      <a class="card" href="<%=H(AppRoot())%>playlist.asp?id=<%=rsPub("ID_PLAYLIST")%>">
        <div class="art">
          <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
          <% Else %><%=CoverPlaceholder(rsPub("TITLE"), "")%><% End If %>
        </div>
        <div class="title"><%=H(rsPub("TITLE"))%></div>
        <div class="meta">Por <%=H(rsPub("USERNAME"))%> &middot; <%=ToInt(rsPub("N"), 0)%> canciones</div>
      </a>
      <%
          rsPub.MoveNext
      Loop
      %>
    </div>
  <% End If %>
  <% rsPub.Close %>
</div>

<!-- ===================== Artistas ===================== -->
<div class="section">
  <div class="section-head">
    <h2>Artistas</h2>
    <a class="more" href="<%=H(AppRoot())%>artists.asp">Ver todos</a>
  </div>
  <% If rsArt.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("mic", 40)%></span><h3>Sin artistas</h3></div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsArt.EOF
          cov = Resolve(rsArt("ARTIST_IMAGE"), "")
      %>
      <a class="card round" href="<%=H(AppRoot())%>artist.asp?id=<%=rsArt("ID_ARTIST")%>">
        <div class="art">
          <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
          <% Else %><%=CoverPlaceholder(rsArt("ARTIST_NAME"), "")%><% End If %>
        </div>
        <div class="title"><%=H(rsArt("ARTIST_NAME") & " " & rsArt("ARTIST_LASTNAME"))%></div>
        <div class="meta">Artista &middot; <%=ToInt(rsArt("N"), 0)%> canciones</div>
      </a>
      <%
          rsArt.MoveNext
      Loop
      %>
    </div>
  <% End If %>
  <% rsArt.Close %>
</div>

<!-- ===================== Albumes ===================== -->
<div class="section">
  <div class="section-head">
    <h2>Albumes</h2>
    <a class="more" href="<%=H(AppRoot())%>albums.asp">Ver todos</a>
  </div>
  <% If rsAlb.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("disc", 40)%></span><h3>Sin albumes</h3></div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsAlb.EOF
          cov = Resolve(rsAlb("ALBUM_IMAGE_PATH"), "")
      %>
      <a class="card" href="<%=H(AppRoot())%>album.asp?id=<%=rsAlb("ID_ALBUM")%>">
        <div class="art">
          <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
          <% Else %><%=CoverPlaceholder(rsAlb("TITLE_ALBUM"), "")%><% End If %>
        </div>
        <div class="title"><%=H(rsAlb("TITLE_ALBUM"))%></div>
        <div class="meta"><%=H(rsAlb("RELEASE_YEAR"))%> &middot; <%=ToInt(rsAlb("N"), 0)%> canciones</div>
      </a>
      <%
          rsAlb.MoveNext
      Loop
      %>
    </div>
  <% End If %>
  <% rsAlb.Close %>
</div>

<!--#include file="includes/footer.asp"-->
