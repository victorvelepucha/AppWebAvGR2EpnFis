<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim id, rsA, rsS, rsAl, cov, pi, url, art, nombre

id = ToInt(Request.QueryString("id"), 0)
If id <= 0 Then GoToUrl AppRoot() & "artists.asp"

Set rsA = Q("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE, " & _
            "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) AS N " & _
            "FROM TBL_ARTIST A WHERE A.ID_ARTIST = ?", Array("i:" & id))
If rsA.EOF Then
    rsA.Close
    GoToUrl AppRoot() & "artists.asp?err=" & Server.URLEncode("El artista solicitado no existe.")
End If

nombre = Trim(CStr(rsA("ARTIST_NAME") & "") & " " & CStr(rsA("ARTIST_LASTNAME") & ""))
cov    = Resolve(rsA("ARTIST_IMAGE"), "")

Set rsS = Q(SqlSongs() & "INNER JOIN TBL_PLAYER P ON P.ID_SONG = S.ID_SONG WHERE P.ID_ARTIST = ? " & _
            "ORDER BY S.PLAYS DESC, S.SONG_NAME ASC", Array("i:" & id))

Set rsAl = Q("SELECT DISTINCT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
             "(SELECT COUNT(*) FROM TBL_TRACKLIST T2 WHERE T2.ID_ALBUM = A.ID_ALBUM) AS N " & _
             "FROM TBL_ALBUM A INNER JOIN TBL_TRACKLIST T ON T.ID_ALBUM = A.ID_ALBUM " & _
             "INNER JOIN TBL_PLAYER P ON P.ID_SONG = T.ID_SONG WHERE P.ID_ARTIST = ? " & _
             "ORDER BY A.RELEASE_YEAR DESC", Array("i:" & id))

PageTitle  = nombre
PageId     = "artists"
NoPlayer   = False
%>
<!--#include file="includes/header.asp"-->

<div class="hero">
  <div class="hero-art round">
    <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="">
    <% Else %><%=CoverPlaceholder(rsA("ARTIST_NAME"), "")%><% End If %>
  </div>
  <div class="hero-info">
    <div class="kicker">Artista verificado</div>
    <h1><%=H(nombre)%></h1>
    <div class="hero-meta">
      <b><%=ToInt(rsA("N"), 0)%></b> canciones en el catalogo
    </div>
    <div class="hero-actions">
      <button type="button" class="big-play" id="btnPlayAll" title="Reproducir todo"><%=Ico("play", 20)%></button>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>artists.asp"><%=Ico("back", 16)%> Volver a artistas</a>
    </div>
  </div>
</div>

<div class="section">
  <div class="section-head"><h2>Canciones de <%=H(nombre)%></h2></div>
  <% If rsS.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("music", 40)%></span><h3>Sin canciones asociadas</h3></div>
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
            rsS.MoveNext
        Loop
        %>
        </tbody>
      </table>
    </div>
  <% End If %>
  <% rsS.Close %>
</div>

<div class="section">
  <div class="section-head"><h2>Albumes donde participa</h2></div>
  <% If rsAl.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("disc", 40)%></span><h3>Sin albumes</h3><p>Este artista no tiene canciones incluidas en albumes.</p></div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsAl.EOF
          url = Resolve(rsAl("ALBUM_IMAGE_PATH"), "")
      %>
      <a class="card" href="<%=H(AppRoot())%>album.asp?id=<%=rsAl("ID_ALBUM")%>">
        <div class="art">
          <% If Len(url) > 0 Then %><img src="<%=H(url)%>" alt="" loading="lazy">
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

<%
rsA.Close
%>
<!--#include file="includes/footer.asp"-->
