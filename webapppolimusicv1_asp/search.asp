<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim qbuscar, likePat, rsS, rsA, rsAl, rsP, cov, pi, url, nS, nA, nAl, nP, total

qbuscar = Trim(CStr(Request.QueryString("q") & ""))
nS = 0 : nA = 0 : nAl = 0 : nP = 0

If Len(qbuscar) > 0 Then
    likePat = "%" & qbuscar & "%"

    Set rsS = Q(SqlSongs() & "WHERE S.SONG_NAME LIKE ? ORDER BY S.PLAYS DESC, S.SONG_NAME ASC", Array("s:" & likePat))

    Set rsA = Q("SELECT TOP 20 A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE, " & _
                "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) AS N " & _
                "FROM TBL_ARTIST A WHERE A.ARTIST_NAME LIKE ? OR A.ARTIST_LASTNAME LIKE ? " & _
                "ORDER BY A.ARTIST_NAME ASC", Array("s:" & likePat, "s:" & likePat))

    Set rsAl = Q("SELECT TOP 20 A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
                 "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
                 "FROM TBL_ALBUM A WHERE A.TITLE_ALBUM LIKE ? OR CAST(A.RELEASE_YEAR AS VARCHAR(10)) LIKE ? " & _
                 "ORDER BY A.RELEASE_YEAR DESC", Array("s:" & likePat, "s:" & likePat))

    Set rsP = Q("SELECT TOP 20 P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, U.USERNAME, " & _
                "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
                "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
                "WHERE P.TITLE LIKE ? AND (P.TYPE = ? OR P.ID_USER = ?) " & _
                "ORDER BY P.CREATION_DATE DESC", _
                Array("s:" & likePat, "i:" & PLAYLIST_PUBLIC, "i:" & CurrentUserID()))
End If

PageTitle  = "Buscar"
PageId     = "search"
PageKicker = "Resultados"
If Len(qbuscar) = 0 Then
    PageSub = "Escribe algo en el buscador de la barra superior para encontrar canciones, artistas, albumes o playlists."
Else
    PageSub = "Resultados para <b>&laquo;" & H(qbuscar) & "&raquo;</b>"
End If
%>
<!--#include file="includes/header.asp"-->

<% If Len(qbuscar) = 0 Then %>

  <div class="empty">
    <span class="ico"><%=Ico("search", 40)%></span>
    <h3>Que quieres escuchar hoy?</h3>
    <p>Usa el buscador superior o explora el catalogo completo.</p>
    <div style="display:flex;gap:10px;justify-content:center;margin-top:18px;flex-wrap:wrap">
      <a class="btn btn-primary" href="<%=H(AppRoot())%>songs.asp"><%=Ico("music", 16)%> Canciones</a>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>artists.asp"><%=Ico("mic", 16)%> Artistas</a>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>albums.asp"><%=Ico("disc", 16)%> Albumes</a>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>playlists.asp"><%=Ico("list", 16)%> Playlists</a>
    </div>
  </div>

<% Else %>

  <!-- ============================ Canciones ============================ -->
  <div class="section">
    <div class="section-head"><h2>Canciones</h2></div>
    <% If rsS.EOF Then %>
      <p class="form-hint">No se encontraron canciones.</p>
    <% Else %>
      <div class="table-wrap">
        <table class="tracklist">
          <thead><tr><th class="idx">#</th><th class="t-cover"></th><th>Cancion</th><th>Reproducciones</th><th class="t-actions"></th></tr></thead>
          <tbody>
          <%
          pi = 0
          Do While Not rsS.EOF
              url = Resolve(rsS("SONG_PATH"), "")
              Response.Write SongRow(pi, rsS("ID_SONG"), rsS("SONG_NAME"), CStr(rsS("ARTISTAS") & ""), rsS("PLAYS"), url, "", "")
              If Len(url) > 0 Then pi = pi + 1
              nS = nS + 1
              rsS.MoveNext
          Loop
          %>
          </tbody>
        </table>
      </div>
    <% End If %>
    <% rsS.Close %>
  </div>

  <!-- ============================ Artistas ============================ -->
  <div class="section">
    <div class="section-head"><h2>Artistas</h2></div>
    <% If rsA.EOF Then %>
      <p class="form-hint">No se encontraron artistas.</p>
    <% Else %>
      <div class="grid">
        <%
        Do While Not rsA.EOF
            cov = Resolve(rsA("ARTIST_IMAGE"), "")
            nA = nA + 1
        %>
        <a class="card round" href="<%=H(AppRoot())%>artist.asp?id=<%=rsA("ID_ARTIST")%>">
          <div class="art">
            <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
            <% Else %><%=CoverPlaceholder(rsA("ARTIST_NAME"), "")%><% End If %>
          </div>
          <div class="title"><%=H(rsA("ARTIST_NAME") & " " & rsA("ARTIST_LASTNAME"))%></div>
          <div class="meta">Artista &middot; <%=ToInt(rsA("N"), 0)%> canciones</div>
        </a>
        <%
            rsA.MoveNext
        Loop
        %>
      </div>
    <% End If %>
    <% rsA.Close %>
  </div>

  <!-- ============================ Albumes ============================ -->
  <div class="section">
    <div class="section-head"><h2>Albumes</h2></div>
    <% If rsAl.EOF Then %>
      <p class="form-hint">No se encontraron albumes.</p>
    <% Else %>
      <div class="grid">
        <%
        Do While Not rsAl.EOF
            cov = Resolve(rsAl("ALBUM_IMAGE_PATH"), "")
            nAl = nAl + 1
        %>
        <a class="card" href="<%=H(AppRoot())%>album.asp?id=<%=rsAl("ID_ALBUM")%>">
          <div class="art">
            <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
            <% Else %><%=CoverPlaceholder(rsAl("TITLE_ALBUM"), "")%><% End If %>
          </div>
          <div class="title"><%=H(rsAl("TITLE_ALBUM"))%></div>
          <div class="meta"><%=H(rsAl("RELEASE_YEAR"))%> &middot; <%=ToInt(rsAl("N"), 0)%> canciones</div>
        </a>
        <%
            rsAl.MoveNext
        Loop
        %>
      </div>
    <% End If %>
    <% rsAl.Close %>
  </div>

  <!-- ============================ Playlists ============================ -->
  <div class="section">
    <div class="section-head"><h2>Playlists</h2></div>
    <% If rsP.EOF Then %>
      <p class="form-hint">No se encontraron playlists.</p>
    <% Else %>
      <div class="grid">
        <%
        Do While Not rsP.EOF
            cov = Resolve(rsP("PHOTO"), "")
            nP = nP + 1
        %>
        <a class="card" href="<%=H(AppRoot())%>playlist.asp?id=<%=rsP("ID_PLAYLIST")%>">
          <div class="art">
            <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
            <% Else %><%=CoverPlaceholder(rsP("TITLE"), "")%><% End If %>
          </div>
          <div class="title"><%=H(rsP("TITLE"))%></div>
          <div class="meta">Por <%=H(rsP("USERNAME"))%> &middot; <%=ToInt(rsP("N"), 0)%> canciones</div>
        </a>
        <%
            rsP.MoveNext
        Loop
        %>
      </div>
    <% End If %>
    <% rsP.Close %>
  </div>

  <% total = nS + nA + nAl + nP %>
  <% If total = 0 Then %>
    <div class="empty" style="margin-top:20px">
      <span class="ico"><%=Ico("search", 40)%></span>
      <h3>Sin resultados</h3>
      <p>No encontramos nada para <b><%=H(qbuscar)%></b>. Prueba con otra palabra.</p>
    </div>
  <% End If %>

<% End If %>

<!--#include file="includes/footer.asp"-->
