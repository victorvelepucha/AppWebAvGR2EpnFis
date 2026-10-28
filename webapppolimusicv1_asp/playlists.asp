<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim scope, rsMine, rsPub, cov, mineCount, pubCount, totalSongs

scope = LCase(Trim(CStr(Request.QueryString("scope") & "")))
If Len(scope) = 0 Then scope = "mine"

mineCount  = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array("i:" & CurrentUserID())), 0)
pubCount   = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE TYPE = ?", Array("i:" & PLAYLIST_PUBLIC)), 0)
totalSongs = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG", Empty), 0)

Set rsMine = Q("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.CREATION_DATE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
               "FROM TBL_PLAYLIST P WHERE P.ID_USER = ? ORDER BY P.CREATION_DATE DESC, P.ID_PLAYLIST DESC", _
               Array("i:" & CurrentUserID()))

Set rsPub = Q("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.CREATION_DATE, U.USERNAME, " & _
              "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
              "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
              "WHERE P.TYPE = ? ORDER BY P.CREATION_DATE DESC, P.ID_PLAYLIST DESC", _
              Array("i:" & PLAYLIST_PUBLIC))

PageTitle  = "Playlists"
PageId     = "playlists"
PageKicker = "Colecciones"
PageSub    = "Tienes <b>" & mineCount & "</b> playlists propias y hay <b>" & pubCount & "</b> publicas en la comunidad."
%>
<!--#include file="includes/header.asp"-->

<div class="hero">
  <div class="hero-art" style="display:grid;place-items:center;background:linear-gradient(135deg,#7a3ff2,#2b0f6b)">
    <div style="font-size:64px;line-height:1"><%=Ico("list", 64)%></div>
  </div>
  <div class="hero-info">
    <div class="kicker">Coleccion</div>
    <h1>Playlists</h1>
    <div class="hero-meta"><b><%=mineCount + pubCount%></b> playlists &middot; <b><%=totalSongs%></b> canciones agregadas</div>
    <div class="hero-actions">
      <a class="btn btn-primary" href="<%=H(AppRoot())%>playlist_edit.asp"><%=Ico("plus", 16)%> Crear playlist</a>
    </div>
  </div>
</div>

<!-- ============================ Mis playlists ============================ -->
<div class="section">
  <div class="section-head">
    <h2>Mis playlists</h2>
    <a class="more" href="<%=H(AppRoot())%>songs.asp">Explorar canciones</a>
  </div>
  <% If rsMine.EOF Then %>
    <div class="empty">
      <span class="ico"><%=Ico("list", 40)%></span>
      <h3>Todavia no tienes playlists</h3>
      <p>Crea una playlist y empieza a agregar tus canciones favoritas.</p>
      <p style="margin-top:14px"><a class="btn btn-primary" href="<%=H(AppRoot())%>playlist_edit.asp"><%=Ico("plus", 16)%> Crear mi primera playlist</a></p>
    </div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsMine.EOF
          cov = Resolve(rsMine("PHOTO"), "")
      %>
      <div class="card">
        <a href="<%=H(AppRoot())%>playlist.asp?id=<%=rsMine("ID_PLAYLIST")%>">
          <div class="art">
            <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
            <% Else %><%=CoverPlaceholder(rsMine("TITLE"), "")%><% End If %>
          </div>
          <div class="title"><%=H(rsMine("TITLE"))%></div>
          <div class="meta">
            <%=ToInt(rsMine("N"), 0)%> canciones &middot;
            <% If ToInt(rsMine("TYPE"), 0) = PLAYLIST_PRIVATE Then %>Privada<% Else %>Publica<% End If %>
          </div>
        </a>
        <div style="display:flex;gap:8px;margin-top:12px">
          <a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>playlist_edit.asp?id=<%=rsMine("ID_PLAYLIST")%>"><%=Ico("edit", 14)%> Editar</a>
          <a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>playlist.asp?id=<%=rsMine("ID_PLAYLIST")%>">Abrir</a>
        </div>
      </div>
      <%
          rsMine.MoveNext
      Loop
      %>
    </div>
  <% End If %>
  <% rsMine.Close %>
</div>

<!-- ============================ Playlists publicas ============================ -->
<div class="section">
  <div class="section-head">
    <h2>Playlists publicas</h2>
    <a class="more" href="<%=H(AppRoot())%>playlists.asp?scope=public">Ver solo publicas</a>
  </div>
  <% If rsPub.EOF Then %>
    <div class="empty"><span class="ico"><%=Ico("list", 40)%></span><h3>Sin playlists publicas</h3></div>
  <% Else %>
    <div class="grid">
      <%
      Do While Not rsPub.EOF
          cov = Resolve(rsPub("PHOTO"), "")
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

<!--#include file="includes/footer.asp"-->
