<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim id, rsP, rsS, rsAvail, cov, pi, url, isOwner, canEdit, act

id = ToInt(Request.QueryString("id"), 0)
If id <= 0 Then GoToUrl AppRoot() & "playlists.asp"

Set rsP = Q("SELECT P.ID_PLAYLIST, P.ID_USER, P.TITLE, P.TYPE, P.PHOTO, P.CREATION_DATE, U.USERNAME " & _
            "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER WHERE P.ID_PLAYLIST = ?", _
            Array("i:" & id))
If rsP.EOF Then
    rsP.Close
    GoToUrl AppRoot() & "playlists.asp?err=" & Server.URLEncode("La playlist solicitada no existe.")
End If

isOwner = (ToInt(rsP("ID_USER"), 0) = CurrentUserID())
canEdit = isOwner Or IsAdmin()

If ToInt(rsP("TYPE"), 0) = PLAYLIST_PRIVATE And Not canEdit Then
    rsP.Close
    GoToUrl AppRoot() & "error.asp?code=403&msg=" & Server.URLEncode("Esta playlist es privada y pertenece a otro usuario.")
End If

cov = Resolve(rsP("PHOTO"), "")

Set rsS = Q(SqlSongs() & "INNER JOIN TBL_PLAYLIST_SONG PS ON PS.ID_SONG = S.ID_SONG " & _
            "WHERE PS.ID_PLAYLIST = ? ORDER BY S.SONG_NAME ASC", Array("i:" & id))

' Canciones que aun no estan en la playlist
Set rsAvail = Q(SqlSongs() & "WHERE NOT EXISTS (SELECT 1 FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = ? AND PS.ID_SONG = S.ID_SONG) " & _
                "ORDER BY S.SONG_NAME ASC", Array("i:" & id))

PageTitle  = CStr(rsP("TITLE") & "")
PageId     = "playlists"
%>
<!--#include file="includes/header.asp"-->

<div class="hero">
  <div class="hero-art">
    <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="">
    <% Else %><%=CoverPlaceholder(rsP("TITLE"), "")%><% End If %>
  </div>
  <div class="hero-info">
    <div class="kicker">
      Playlist <% If ToInt(rsP("TYPE"), 0) = PLAYLIST_PRIVATE Then %>privada<% Else %>publica<% End If %>
    </div>
    <h1><%=H(rsP("TITLE"))%></h1>
    <div class="hero-meta">
      Creada por <b><%=H(rsP("USERNAME"))%></b> &middot;
      <%=FmtDate(rsP("CREATION_DATE"))%> &middot;
      <span id="plCount">canciones</span>
    </div>
    <div class="hero-actions">
      <button type="button" class="big-play" id="btnPlayAll" title="Reproducir playlist"><%=Ico("play", 20)%></button>
      <a class="btn btn-outline" href="<%=H(AppRoot())%>playlists.asp"><%=Ico("back", 16)%> Volver</a>
      <% If canEdit Then %>
        <a class="btn btn-outline" href="<%=H(AppRoot())%>playlist_edit.asp?id=<%=id%>"><%=Ico("edit", 16)%> Editar</a>
      <% End If %>
    </div>
  </div>
</div>

<% If canEdit And Not rsAvail.EOF Then %>
<div class="section" style="background:#181818;border:1px solid #242424;border-radius:12px;padding:16px 18px">
  <form method="post" action="<%=H(AppRoot())%>playlist_item.asp" style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
    <input type="hidden" name="action" value="add">
    <input type="hidden" name="playlist" value="<%=id%>">
    <div style="flex:1 1 320px">
      <label style="display:block;font-size:11.5px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#a7a7a7;margin-bottom:6px">
        Agregar cancion a esta playlist</label>
      <select class="input" name="song" required>
        <%
        Do While Not rsAvail.EOF
        %>
          <option value="<%=rsAvail("ID_SONG")%>"><%=H(rsAvail("SONG_NAME"))%> &mdash; <%=H(Shorten(rsAvail("ARTISTAS"), 40))%></option>
        <%
            rsAvail.MoveNext
        Loop
        %>
      </select>
    </div>
    <button type="submit" class="btn btn-primary"><%=Ico("plus", 16)%> Agregar</button>
  </form>
</div>
<% End If %>
<% rsAvail.Close %>

<div class="section">
  <div class="section-head"><h2>Canciones de la playlist</h2></div>
  <% If rsS.EOF Then %>
    <div class="empty">
      <span class="ico"><%=Ico("music", 40)%></span>
      <h3>Esta playlist esta vacia</h3>
      <p>Agrega canciones usando el selector de arriba o desde la seccion <a href="<%=H(AppRoot())%>songs.asp">Canciones</a>.</p>
    </div>
  <% Else %>
    <div class="table-wrap">
      <table class="tracklist">
        <thead><tr><th class="idx">#</th><th class="t-cover"></th><th>Cancion</th><th>Reproducciones</th><th class="t-actions">Acciones</th></tr></thead>
        <tbody>
        <%
        pi = 0
        Do While Not rsS.EOF
            url = Resolve(rsS("SONG_PATH"), "")
            act = ""
            If canEdit Then
                act = "<form method=""post"" action=""" & H(AppRoot()) & "playlist_item.asp"" style=""display:inline"" " & _
                      "data-confirm=""Se quitara la cancion de esta playlist. Continuar?"">" & _
                      "<input type=""hidden"" name=""action"" value=""remove"">" & _
                      "<input type=""hidden"" name=""playlist"" value=""" & id & """>" & _
                      "<input type=""hidden"" name=""song"" value=""" & rsS("ID_SONG") & """>" & _
                      "<button type=""submit"" class=""row-btn danger"" title=""Quitar de la playlist"">" & Ico("x", 12) & "</button></form>"
            End If
            Response.Write SongRow(pi, rsS("ID_SONG"), rsS("SONG_NAME"), CStr(rsS("ARTISTAS") & ""), rsS("PLAYS"), url, "", act)
            If Len(url) > 0 Then pi = pi + 1
            rsS.MoveNext
        Loop
        %>
        </tbody>
      </table>
    </div>
    <p class="form-hint" style="margin-top:12px"><b><%=pi%></b> pista(s) reproducible(s) en esta playlist.</p>
  <% End If %>
  <% rsS.Close %>
</div>

<%
rsP.Close
%>
<!--#include file="includes/footer.asp"-->
