<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim qbuscar, sort, orderBy, page, total, pages, offset, rs, rsPl, pi, url, cov
Dim act, sql, whereSql, params, cntParams

qbuscar    = Trim(CStr(Request.QueryString("q") & ""))
sort = LCase(Trim(CStr(Request.QueryString("sort") & "")))
page = ToInt(Request.QueryString("p"), 1)
If page < 1 Then page = 1

Select Case sort
    Case "name"  : orderBy = "S.SONG_NAME ASC"
    Case "recent": orderBy = "S.ID_SONG DESC"
    Case "az"    : orderBy = "S.SONG_NAME ASC"
    Case Else    : sort = "plays" : orderBy = "S.PLAYS DESC, S.SONG_NAME ASC"
End Select

whereSql = ""
params = Empty
cntParams = Empty
If Len(qbuscar) > 0 Then
    whereSql = "WHERE S.SONG_NAME LIKE ? "
    params = Array("s:%" & qbuscar & "%")
    cntParams = Array("s:%" & qbuscar & "%")
End If

total = ToInt(XS("SELECT COUNT(*) FROM TBL_SONG S " & whereSql, cntParams), 0)
pages = Int((total + ROWS_PER_PAGE - 1) / ROWS_PER_PAGE)
If pages < 1 Then pages = 1
If page > pages Then page = pages
offset = (page - 1) * ROWS_PER_PAGE

sql = SqlSongs() & whereSql & "ORDER BY " & orderBy & " OFFSET " & offset & " ROWS FETCH NEXT " & ROWS_PER_PAGE & " ROWS ONLY"
Set rs = Q(sql, params)

Set rsPl = Q("SELECT ID_PLAYLIST, TITLE FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY TITLE ASC", _
             Array("i:" & CurrentUserID()))

PageTitle  = "Canciones"
PageId     = "songs"
PageKicker = "Biblioteca"
PageSub    = "<b>" & total & "</b> canciones disponibles en PoliMusic."
%>
<!--#include file="includes/header.asp"-->

<div class="section">
  <form method="get" action="<%=H(AppRoot())%>songs.asp" class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Filtrar por nombre de cancion"
           style="max-width:300px">
    <select class="input" name="sort" style="max-width:190px">
      <option value="plays"<% If sort = "plays" Then %> selected<% End If %>>Mas escuchadas</option>
      <option value="name"<% If sort = "name" Then %> selected<% End If %>>Nombre (A-Z)</option>
      <option value="recent"<% If sort = "recent" Then %> selected<% End If %>>Mas recientes</option>
    </select>
    <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Aplicar</button>
    <% If Len(qbuscar) > 0 Then %>
      <a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>songs.asp">Limpiar</a>
    <% End If %>
  </form>
</div>

<% If rs.EOF Then %>
  <div class="empty">
    <span class="ico"><%=Ico("music", 40)%></span>
    <h3>No se encontraron canciones</h3>
    <p>Prueba con otro criterio de busqueda.</p>
  </div>
<% Else %>
  <div class="table-wrap">
    <table class="tracklist">
      <thead>
        <tr>
          <th class="idx">#</th>
          <th class="t-cover"></th>
          <th>Cancion</th>
          <th>Reproducciones</th>
          <th class="t-actions">Acciones</th>
        </tr>
      </thead>
      <tbody>
      <%
      pi = 0
      Do While Not rs.EOF
          url = Resolve(rs("SONG_PATH"), "")
          act = ""
          If rsPl.RecordCount > 0 Then
              act = "<form method=""post"" action=""" & H(AppRoot()) & "playlist_item.asp"" style=""display:inline-flex;gap:4px;align-items:center"">" & _
                    "<input type=""hidden"" name=""action"" value=""add"">" & _
                    "<input type=""hidden"" name=""song"" value=""" & rs("ID_SONG") & """>" & _
                    "<input type=""hidden"" name=""back"" value=""" & H("songs.asp?p=" & page & "&q=" & Server.URLEncode(qbuscar) & "&sort=" & sort) & """>" & _
                    "<select name=""playlist"" class=""input"" style=""width:auto;padding:4px 8px;font-size:12px"">"
              rsPl.MoveFirst
              Do While Not rsPl.EOF
                  act = act & "<option value=""" & rsPl("ID_PLAYLIST") & """>" & H(Shorten(rsPl("TITLE"), 18)) & "</option>"
                  rsPl.MoveNext
              Loop
              act = act & "</select><button type=""submit"" class=""row-btn"" title=""Agregar a playlist"">" & Ico("plus", 12) & "</button></form>"
          End If
          Response.Write SongRow(pi, rs("ID_SONG"), rs("SONG_NAME"), CStr(rs("ARTISTAS") & ""), rs("PLAYS"), url, "", act)
          If Len(url) > 0 Then pi = pi + 1
          rs.MoveNext
      Loop
      %>
      </tbody>
    </table>
  </div>

  <% If pages > 1 Then %>
    <div style="display:flex;gap:8px;justify-content:center;margin-top:26px;flex-wrap:wrap">
      <% If page > 1 Then %>
        <a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>songs.asp?p=<%=page-1%>&q=<%=Server.URLEncode(qbuscar)%>&sort=<%=H(sort)%>">&laquo; Anterior</a>
      <% End If %>
      <span class="btn btn-sm" style="background:#242424;cursor:default">Pagina <%=page%> de <%=pages%></span>
      <% If page < pages Then %>
        <a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>songs.asp?p=<%=page+1%>&q=<%=Server.URLEncode(qbuscar)%>&sort=<%=H(sort)%>">Siguiente &raquo;</a>
      <% End If %>
    </div>
  <% End If %>
<% End If %>

<%
rs.Close
rsPl.Close
%>
<!--#include file="includes/footer.asp"-->
