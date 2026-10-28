<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim qbuscar, rs, total, act, id, tipo, scope

qbuscar     = Trim(CStr(Request.QueryString("q") & ""))
scope = LCase(Trim(CStr(Request.QueryString("scope") & "")))
act   = LCase(Trim(CStr(Request.Form("action") & "")))
id    = ToInt(Request.Form("id"), 0)

If act = "delete" And id > 0 Then
    X "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array("i:" & id)
    X "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array("i:" & id)
    SetFlash "Playlist eliminada correctamente.", "ok"
    GoToUrl AppRoot() & "admin/playlists.asp"
End If

If Len(qbuscar) > 0 Then
    Set rs = Q("SELECT P.ID_PLAYLIST, P.TITLE, P.TYPE, P.CREATION_DATE, P.PHOTO, U.USERNAME, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
               "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
               "WHERE P.TITLE LIKE ? ORDER BY P.ID_PLAYLIST ASC", Array("s:%" & qbuscar & "%"))
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE TITLE LIKE ?", Array("s:%" & qbuscar & "%")), 0)
Else
    Set rs = Q("SELECT P.ID_PLAYLIST, P.TITLE, P.TYPE, P.CREATION_DATE, P.PHOTO, U.USERNAME, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS N " & _
               "FROM TBL_PLAYLIST P INNER JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
               "ORDER BY P.ID_PLAYLIST ASC", Empty)
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty), 0)
End If

PageTitle  = "Gestion de playlists"
PageId     = "admin_playlists"
PageKicker = "Administracion"
PageSub    = "<b>" & total & "</b> playlist(s) en el sistema."
%>
<!--#include file="../includes/header.asp"-->

<div class="section">
  <div class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <form method="get" action="<%=H(AppRoot())%>admin/playlists.asp" style="display:flex;gap:10px;align-items:center">
      <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar playlist por titulo" style="max-width:280px">
      <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Buscar</button>
      <% If Len(qbuscar) > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>admin/playlists.asp">Limpiar</a><% End If %>
    </form>
    <a class="btn btn-primary btn-sm" href="<%=H(AppRoot())%>playlist_edit.asp" style="margin-left:auto"><%=Ico("plus", 15)%> Nueva playlist</a>
  </div>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("list", 40)%></span><h3>Sin playlists</h3></div>
<% Else %>
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr><th>ID</th><th>Portada</th><th>Titulo</th><th>Propietario</th><th>Tipo</th><th>Canciones</th><th>Creada</th><th class="actions">Acciones</th></tr>
      </thead>
      <tbody>
      <%
      Do While Not rs.EOF
          tipo = ToInt(rs("TYPE"), 0)
      %>
      <tr>
        <td><%=rs("ID_PLAYLIST")%></td>
        <td>
          <% If Len(Resolve(rs("PHOTO"), "")) > 0 Then %>
            <img class="thumb" src="<%=H(Resolve(rs("PHOTO"), ""))%>" alt="">
          <% Else %>
            <div style="width:38px;height:38px;border-radius:4px;overflow:hidden"><%=CoverPlaceholder(rs("TITLE"), "sm")%></div>
          <% End If %>
        </td>
        <td><b><%=H(rs("TITLE"))%></b></td>
        <td><%=H(rs("USERNAME"))%></td>
        <td>
          <% If tipo = PLAYLIST_PRIVATE Then %>
            <span class="tag tag-gray">Privada (1)</span>
          <% Else %>
            <span class="tag tag-blue">Publica (0)</span>
          <% End If %>
        </td>
        <td><%=ToInt(rs("N"), 0)%></td>
        <td><%=FmtDate(rs("CREATION_DATE"))%></td>
        <td class="actions">
          <a class="row-btn" title="Ver" href="<%=H(AppRoot())%>playlist.asp?id=<%=rs("ID_PLAYLIST")%>"><%=Ico("search", 12)%></a>
          <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>playlist_edit.asp?id=<%=rs("ID_PLAYLIST")%>"><%=Ico("edit", 12)%></a>
          <form method="post" action="<%=H(AppRoot())%>admin/playlists.asp" style="display:inline"
                data-confirm="Se eliminara la playlist <%=H(rs("TITLE"))%> y su contenido. Continuar?">
            <input type="hidden" name="action" value="delete">
            <input type="hidden" name="id" value="<%=rs("ID_PLAYLIST")%>">
            <button type="submit" class="row-btn danger" title="Eliminar"><%=Ico("trash", 12)%></button>
          </form>
        </td>
      </tr>
      <%
          rs.MoveNext
      Loop
      %>
      </tbody>
    </table>
  </div>
<% End If %>

<%
rs.Close
%>
<!--#include file="../includes/footer.asp"-->
