<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim qbuscar, rs, total, act, id, anio

qbuscar    = Trim(CStr(Request.QueryString("q") & ""))
anio = ToInt(Request.QueryString("year"), 0)
act  = LCase(Trim(CStr(Request.Form("action") & "")))
id   = ToInt(Request.Form("id"), 0)

If act = "delete" And id > 0 Then
    X "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array("i:" & id)
    X "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array("i:" & id)
    SetFlash "Album eliminado junto con su lista de canciones.", "ok"
    GoToUrl AppRoot() & "admin/albums.asp"
End If

If Len(qbuscar) > 0 Then
    Set rs = Q("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
               "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
               "FROM TBL_ALBUM A WHERE A.TITLE_ALBUM LIKE ? ORDER BY A.ID_ALBUM ASC", Array("s:%" & qbuscar & "%"))
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_ALBUM WHERE TITLE_ALBUM LIKE ?", Array("s:%" & qbuscar & "%")), 0)
Else
    Set rs = Q("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH, " & _
               "(SELECT COUNT(*) FROM TBL_TRACKLIST T WHERE T.ID_ALBUM = A.ID_ALBUM) AS N " & _
               "FROM TBL_ALBUM A ORDER BY A.ID_ALBUM ASC", Empty)
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_ALBUM", Empty), 0)
End If

PageTitle  = "Gestion de albumes"
PageId     = "admin_albums"
PageKicker = "Administracion"
PageSub    = "<b>" & total & "</b> album(es) registrado(s)."
%>
<!--#include file="../includes/header.asp"-->

<div class="section">
  <div class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <form method="get" action="<%=H(AppRoot())%>admin/albums.asp" style="display:flex;gap:10px;align-items:center">
      <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar album" style="max-width:280px">
      <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Buscar</button>
      <% If Len(qbuscar) > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>admin/albums.asp">Limpiar</a><% End If %>
    </form>
    <a class="btn btn-primary btn-sm" href="<%=H(AppRoot())%>admin/album_form.asp" style="margin-left:auto"><%=Ico("plus", 15)%> Nuevo album</a>
  </div>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("disc", 40)%></span><h3>Sin albumes</h3></div>
<% Else %>
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr><th>ID</th><th>Portada</th><th>Titulo</th><th>Anio</th><th>Imagen (ruta)</th><th>Canciones</th><th class="actions">Acciones</th></tr>
      </thead>
      <tbody>
      <%
      Do While Not rs.EOF
      %>
      <tr>
        <td><%=rs("ID_ALBUM")%></td>
        <td>
          <% If Len(Resolve(rs("ALBUM_IMAGE_PATH"), "")) > 0 Then %>
            <img class="thumb" src="<%=H(Resolve(rs("ALBUM_IMAGE_PATH"), ""))%>" alt="">
          <% Else %>
            <div style="width:38px;height:38px;border-radius:4px;overflow:hidden"><%=CoverPlaceholder(rs("TITLE_ALBUM"), "sm")%></div>
          <% End If %>
        </td>
        <td><b><%=H(rs("TITLE_ALBUM"))%></b></td>
        <td><%=H(rs("RELEASE_YEAR"))%></td>
        <td><code style="font-size:11.5px;color:#a7a7a7"><%=H(rs("ALBUM_IMAGE_PATH"))%></code></td>
        <td><%=ToInt(rs("N"), 0)%></td>
        <td class="actions">
          <a class="row-btn" title="Ver" href="<%=H(AppRoot())%>album.asp?id=<%=rs("ID_ALBUM")%>"><%=Ico("search", 12)%></a>
          <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>admin/album_form.asp?id=<%=rs("ID_ALBUM")%>"><%=Ico("edit", 12)%></a>
          <form method="post" action="<%=H(AppRoot())%>admin/albums.asp" style="display:inline"
                data-confirm="Se eliminara el album <%=H(rs("TITLE_ALBUM"))%> y su tracklist. Continuar?">
            <input type="hidden" name="action" value="delete">
            <input type="hidden" name="id" value="<%=rs("ID_ALBUM")%>">
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
