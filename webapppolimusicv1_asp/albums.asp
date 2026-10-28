<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim qbuscar, rs, cov, total, anio, orderBy

qbuscar    = Trim(CStr(Request.QueryString("q") & ""))
anio = ToInt(Request.QueryString("year"), 0)
orderBy = "A.RELEASE_YEAR DESC, A.TITLE_ALBUM ASC"

If Len(qbuscar) > 0 Then
    Set rs = Q("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
               "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
               "FROM TBL_ALBUM A WHERE A.TITLE_ALBUM LIKE ? ORDER BY " & orderBy, Array("s:%" & qbuscar & "%"))
ElseIf anio > 0 Then
    Set rs = Q("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
               "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
               "FROM TBL_ALBUM A WHERE A.RELEASE_YEAR = ? ORDER BY " & orderBy, Array("i:" & anio))
Else
    Set rs = Q("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
               "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
               "FROM TBL_ALBUM A ORDER BY " & orderBy, Empty)
End If

total = ToInt(XS("SELECT COUNT(*) FROM TBL_ALBUM", Empty), 0)

PageTitle  = "Albumes"
PageId     = "albums"
PageKicker = "Biblioteca"
PageSub    = "<b>" & total & "</b> albumes en el catalogo."
%>
<!--#include file="includes/header.asp"-->

<div class="section">
  <form method="get" action="<%=H(AppRoot())%>albums.asp" class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar album por titulo" style="max-width:300px">
    <input class="input" type="number" name="year" value="<% If anio > 0 Then %><%=anio%><% End If %>" placeholder="Anio" style="max-width:120px">
    <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Filtrar</button>
    <% If Len(qbuscar) > 0 Or anio > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>albums.asp">Limpiar</a><% End If %>
  </form>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("disc", 40)%></span><h3>Sin albumes</h3><p>No hay albumes que coincidan con el filtro aplicado.</p></div>
<% Else %>
  <div class="grid grid-lg">
    <%
    Do While Not rs.EOF
        cov = Resolve(rs("ALBUM_IMAGE_PATH"), "")
    %>
    <a class="card" href="<%=H(AppRoot())%>album.asp?id=<%=rs("ID_ALBUM")%>">
      <div class="art">
        <% If Len(cov) > 0 Then %><img src="<%=H(cov)%>" alt="" loading="lazy">
        <% Else %><%=CoverPlaceholder(rs("TITLE_ALBUM"), "")%><% End If %>
      </div>
      <div class="title"><%=H(rs("TITLE_ALBUM"))%></div>
      <div class="meta">
        <%=H(rs("RELEASE_YEAR"))%> &middot; <%=ToInt(rs("N"), 0)%> canciones
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
