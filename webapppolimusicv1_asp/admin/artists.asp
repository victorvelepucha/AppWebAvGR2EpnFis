<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim qbuscar, rs, total, act, id, nSongs

qbuscar   = Trim(CStr(Request.QueryString("q") & ""))
act = LCase(Trim(CStr(Request.Form("action") & "")))
id  = ToInt(Request.Form("id"), 0)

If act = "delete" And id > 0 Then
    X "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array("i:" & id)
    X "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array("i:" & id)
    SetFlash "Artista eliminado y desvinculado de sus canciones.", "ok"
    GoToUrl AppRoot() & "admin/artists.asp"
End If

If Len(qbuscar) > 0 Then
    Set rs = Q("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) AS N " & _
               "FROM TBL_ARTIST A WHERE A.ARTIST_NAME LIKE ? OR A.ARTIST_LASTNAME LIKE ? ORDER BY A.ID_ARTIST ASC", _
               Array("s:%" & qbuscar & "%", "s:%" & qbuscar & "%"))
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_ARTIST WHERE ARTIST_NAME LIKE ? OR ARTIST_LASTNAME LIKE ?", _
                     Array("s:%" & qbuscar & "%", "s:%" & qbuscar & "%")), 0)
Else
    Set rs = Q("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) AS N " & _
               "FROM TBL_ARTIST A ORDER BY A.ID_ARTIST ASC", Empty)
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_ARTIST", Empty), 0)
End If

PageTitle  = "Gestion de artistas"
PageId     = "admin_artists"
PageKicker = "Administracion"
PageSub    = "<b>" & total & "</b> artista(s) registrado(s)."
%>
<!--#include file="../includes/header.asp"-->

<div class="section">
  <div class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <form method="get" action="<%=H(AppRoot())%>admin/artists.asp" style="display:flex;gap:10px;align-items:center">
      <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar artista" style="max-width:280px">
      <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Buscar</button>
      <% If Len(qbuscar) > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>admin/artists.asp">Limpiar</a><% End If %>
    </form>
    <a class="btn btn-primary btn-sm" href="<%=H(AppRoot())%>admin/artist_form.asp" style="margin-left:auto"><%=Ico("plus", 15)%> Nuevo artista</a>
  </div>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("mic", 40)%></span><h3>Sin artistas</h3></div>
<% Else %>
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr><th>ID</th><th>Imagen</th><th>Nombre</th><th>Apellido</th><th>Imagen (ruta)</th><th>Canciones</th><th class="actions">Acciones</th></tr>
      </thead>
      <tbody>
      <%
      Do While Not rs.EOF
      %>
      <tr>
        <td><%=rs("ID_ARTIST")%></td>
        <td>
          <% If Len(Resolve(rs("ARTIST_IMAGE"), "")) > 0 Then %>
            <img class="thumb round" src="<%=H(Resolve(rs("ARTIST_IMAGE"), ""))%>" alt="">
          <% Else %>
            <div style="width:38px;height:38px;border-radius:50%;overflow:hidden"><%=CoverPlaceholder(rs("ARTIST_NAME"), "sm")%></div>
          <% End If %>
        </td>
        <td><b><%=H(rs("ARTIST_NAME"))%></b></td>
        <td><%=H(rs("ARTIST_LASTNAME"))%></td>
        <td><code style="font-size:11.5px;color:#a7a7a7"><%=H(rs("ARTIST_IMAGE"))%></code></td>
        <td><%=ToInt(rs("N"), 0)%></td>
        <td class="actions">
          <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>admin/artist_form.asp?id=<%=rs("ID_ARTIST")%>"><%=Ico("edit", 12)%></a>
          <form method="post" action="<%=H(AppRoot())%>admin/artists.asp" style="display:inline"
                data-confirm="Se eliminara el artista <%=H(rs("ARTIST_NAME"))%> y se desvinculara de sus canciones. Continuar?">
            <input type="hidden" name="action" value="delete">
            <input type="hidden" name="id" value="<%=rs("ID_ARTIST")%>">
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
