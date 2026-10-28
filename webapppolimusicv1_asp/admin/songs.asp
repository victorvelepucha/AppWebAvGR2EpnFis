<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim qbuscar, rs, total, act, id, msg

qbuscar   = Trim(CStr(Request.QueryString("q") & ""))
act = LCase(Trim(CStr(Request.Form("action") & "")))
id  = ToInt(Request.Form("id"), 0)

If act = "delete" And id > 0 Then
    X "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array("i:" & id)
    X "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array("i:" & id)
    X "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array("i:" & id)
    X "DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array("i:" & id)
    SetFlash "Cancion eliminada junto con sus relaciones (artistas, albumes y playlists).", "ok"
    GoToUrl AppRoot() & "admin/songs.asp"
End If

If Len(qbuscar) > 0 Then
    Set rs = Q(SqlSongs() & "WHERE S.SONG_NAME LIKE ? ORDER BY S.ID_SONG ASC", Array("s:%" & qbuscar & "%"))
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_SONG WHERE SONG_NAME LIKE ?", Array("s:%" & qbuscar & "%")), 0)
Else
    Set rs = Q(SqlSongs() & "ORDER BY S.ID_SONG ASC", Empty)
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_SONG", Empty), 0)
End If

PageTitle  = "Gestion de canciones"
PageId     = "admin_songs"
PageKicker = "Administracion"
PageSub    = "<b>" & total & "</b> cancion(es) en el catalogo."
%>
<!--#include file="../includes/header.asp"-->

<div class="section">
  <div class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <form method="get" action="<%=H(AppRoot())%>admin/songs.asp" style="display:flex;gap:10px;align-items:center">
      <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar cancion" style="max-width:280px">
      <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Buscar</button>
      <% If Len(qbuscar) > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>admin/songs.asp">Limpiar</a><% End If %>
    </form>
    <a class="btn btn-primary btn-sm" href="<%=H(AppRoot())%>admin/song_form.asp" style="margin-left:auto"><%=Ico("plus", 15)%> Nueva cancion</a>
  </div>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("music", 40)%></span><h3>Sin canciones</h3></div>
<% Else %>
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr><th>ID</th><th>Portada</th><th>Cancion</th><th>Artistas</th><th>Ruta del archivo</th><th>Archivo</th><th>Reproducciones</th><th class="actions">Acciones</th></tr>
      </thead>
      <tbody>
      <%
      Do While Not rs.EOF
      %>
      <tr>
        <td><%=rs("ID_SONG")%></td>
        <td>
          <% If SongExists(rs("SONG_PATH")) Then %>
            <span class="tag tag-green"><%=Ico("check", 12)%></span>
          <% Else %>
            <span class="tag tag-red">!</span>
          <% End If %>
        </td>
        <td><b><%=H(rs("SONG_NAME"))%></b></td>
        <td><%=H(rs("ARTISTAS"))%></td>
        <td><code style="font-size:11.5px;color:#a7a7a7"><%=H(rs("SONG_PATH"))%></code></td>
        <td>
          <% If SongExists(rs("SONG_PATH")) Then %>
            <span class="tag tag-green">Disponible</span>
          <% Else %>
            <span class="tag tag-red">No encontrado</span>
          <% End If %>
        </td>
        <td><%=ToInt(rs("PLAYS"), 0)%></td>
        <td class="actions">
          <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>admin/song_form.asp?id=<%=rs("ID_SONG")%>"><%=Ico("edit", 12)%></a>
          <form method="post" action="<%=H(AppRoot())%>admin/songs.asp" style="display:inline"
                data-confirm="Se eliminara la cancion <%=H(rs("SONG_NAME"))%> y se quitara de albumes y playlists. Continuar?">
            <input type="hidden" name="action" value="delete">
            <input type="hidden" name="id" value="<%=rs("ID_SONG")%>">
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
