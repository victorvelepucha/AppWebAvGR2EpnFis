<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim id, rsA, rsT, cov, pi, url, art, totalDur, durSec

id = ToInt(Request.QueryString("id"), 0)
If id <= 0 Then GoToUrl AppRoot() & "albums.asp"

Set rsA = Q("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
            "ISNULL(STUFF((SELECT DISTINCT ', ' + AR.ARTIST_NAME + ISNULL(' ' + AR.ARTIST_LASTNAME, '') " & _
            "FROM TBL_TRACKLIST T INNER JOIN TBL_PLAYER P ON P.ID_SONG = T.ID_SONG " & _
            "INNER JOIN TBL_ARTIST AR ON AR.ID_ARTIST = P.ID_ARTIST " & _
            "WHERE T.ID_ALBUM = A.ID_ALBUM FOR XML PATH('')), 1, 2, ''), 'Varios artistas') AS ARTISTAS " & _
            "FROM TBL_ALBUM A WHERE A.ID_ALBUM = ?", Array("i:" & id))
If rsA.EOF Then
    rsA.Close
    GoToUrl AppRoot() & "albums.asp?err=" & Server.URLEncode("El album solicitado no existe.")
End If

cov = Resolve(rsA("ALBUM_IMAGE_PATH"), "")

Set rsT = Q(SqlSongs() & "INNER JOIN TBL_TRACKLIST T ON T.ID_SONG = S.ID_SONG WHERE T.ID_ALBUM = ? " & _
            "ORDER BY S.ID_SONG ASC", Array("i:" & id))

PageTitle  = CStr(rsA("TITLE_ALBUM") & "")
PageId     = "albums"
%>
<!--#include file="includes/header.asp"-->

<div class="hero">
  <div class="hero-art">
    <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="">
    <% Else %><%=CoverPlaceholder(rsA("TITLE_ALBUM"), "")%><% End If %>
  </div>
  <div class="hero-info">
    <div class="kicker">Album</div>
    <h1><%=H(rsA("TITLE_ALBUM"))%></h1>
    <div class="hero-meta">
      <b><%=H(rsA("ARTISTAS"))%></b> &nbsp;&middot;&nbsp;
      <%=H(rsA("RELEASE_YEAR"))%> &nbsp;&middot;&nbsp;
      <span id="albumCount">canciones</span>
    </div>
    <div class="hero-actions">
      <button type="button" class="big-play" id="btnPlayAll" title="Reproducir album"><%=Ico("play", 20)%></button>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>albums.asp"><%=Ico("back", 16)%> Volver a albumes</a>
      <% If IsAdmin() Then %>
        <a class="btn btn-outline" href="<%=H(AppRoot())%>admin/album_form.asp?id=<%=id%>"><%=Ico("edit", 16)%> Editar album</a>
      <% End If %>
    </div>
  </div>
</div>

<div class="section">
  <div class="section-head"><h2>Lista de canciones</h2></div>
  <% If rsT.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("music", 40)%></span><h3>Album sin canciones</h3><p>Agrega canciones desde el panel de administracion.</p></div>
  <% Else %>
    <div class="table-wrap">
      <table class="tracklist">
        <thead><tr><th class="idx">#</th><th class="t-cover"></th><th>Cancion</th><th>Reproducciones</th><th class="t-actions"></th></tr></thead>
        <tbody>
        <%
        pi = 0
        Do While Not rsT.EOF
            url = Resolve(rsT("SONG_PATH"), "")
            Response.Write SongRow(pi, rsT("ID_SONG"), rsT("SONG_NAME"), CStr(rsT("ARTISTAS") & ""), rsT("PLAYS"), url, "", "")
            If Len(url) > 0 Then pi = pi + 1
            rsT.MoveNext
        Loop
        %>
        </tbody>
      </table>
    </div>
    <p class="form-hint" style="margin-top:12px">Total: <b><%=pi%></b> pista(s) disponible(s) para reproduccion.</p>
  <% End If %>
  <% rsT.Close %>
</div>

<%
rsA.Close
%>
<!--#include file="includes/footer.asp"-->
