<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim qbuscar, rs, cov, total

qbuscar = Trim(CStr(Request.QueryString("q") & ""))

If Len(qbuscar) > 0 Then
    Set rs = Q("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) AS N, " & _
               "(SELECT COUNT(DISTINCT T.ID_ALBUM) FROM TBL_TRACKLIST T INNER JOIN TBL_PLAYER P2 ON P2.ID_SONG = T.ID_SONG WHERE P2.ID_ARTIST = A.ID_ARTIST) AS ALB " & _
               "FROM TBL_ARTIST A " & _
               "WHERE A.ARTIST_NAME LIKE ? OR A.ARTIST_LASTNAME LIKE ? " & _
               "ORDER BY A.ARTIST_NAME ASC", Array("s:%" & qbuscar & "%", "s:%" & qbuscar & "%"))
Else
    Set rs = Q("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) AS N, " & _
               "(SELECT COUNT(DISTINCT T.ID_ALBUM) FROM TBL_TRACKLIST T INNER JOIN TBL_PLAYER P2 ON P2.ID_SONG = T.ID_SONG WHERE P2.ID_ARTIST = A.ID_ARTIST) AS ALB " & _
               "FROM TBL_ARTIST A ORDER BY A.ARTIST_NAME ASC", Empty)
End If

total = ToInt(XS("SELECT COUNT(*) FROM TBL_ARTIST", Empty), 0)

PageTitle  = "Artistas"
PageId     = "artists"
PageKicker = "Biblioteca"
PageSub    = "<b>" & total & "</b> artistas registrados en la plataforma."
%>
<!--#include file="includes/header.asp"-->

<div class="section">
  <form method="get" action="<%=H(AppRoot())%>artists.asp" class="section-head" style="gap:10px;align-items:center">
    <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar artista por nombre o apellido" style="max-width:340px">
    <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Buscar</button>
    <% If Len(qbuscar) > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>artists.asp">Limpiar</a><% End If %>
  </form>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("mic", 40)%></span><h3>Sin resultados</h3><p>No hay artistas que coincidan con la busqueda.</p></div>
<% Else %>
  <div class="grid grid-lg">
    <%
    Do While Not rs.EOF
        cov = Resolve(rs("ARTIST_IMAGE"), "")
    %>
    <a class="card round" href="<%=H(AppRoot())%>artist.asp?id=<%=rs("ID_ARTIST")%>">
      <div class="art">
        <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
        <% Else %><%=CoverPlaceholder(rs("ARTIST_NAME"), "")%><% End If %>
      </div>
      <div class="title"><%=H(rs("ARTIST_NAME") & " " & rs("ARTIST_LASTNAME"))%></div>
      <div class="meta">
        <%=ToInt(rs("N"), 0)%> canciones &middot; <%=ToInt(rs("ALB"), 0)%> albumes
      </div>
    </a>
    <%
        rs.MoveNext
    Loop
    %>
  </div>
<% End If %>

<%
rs.Close
%>
<!--#include file="includes/footer.asp"-->
