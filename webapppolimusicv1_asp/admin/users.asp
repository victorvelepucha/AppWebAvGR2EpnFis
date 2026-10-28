<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/lib.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
RequireAdmin

Dim qbuscar, rs, total, act, id, msg, kind, nPlaylists, isAdminUser

qbuscar = Trim(CStr(Request.QueryString("q") & ""))
act = LCase(Trim(CStr(Request.Form("action") & "")))
id = ToInt(Request.Form("id"), 0)

' ------------------------------------------------------------- Eliminar usuario
If act = "delete" And id > 0 Then
    If id = CurrentUserID() Then
        SetFlash "No puedes eliminar tu propia cuenta mientras la tienes abierta.", "error"
    Else
        isAdminUser = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE ID_USER = ? AND USER_TYPE = ?", _
                               Array("i:" & id, "i:" & USER_TYPE_ADMIN)), 0)
        If isAdminUser > 0 And ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = ?", _
                                        Array("i:" & USER_TYPE_ADMIN)), 0) <= 1 Then
            SetFlash "No se puede eliminar al ultimo administrador del sistema.", "error"
        Else
            nPlaylists = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array("i:" & id)), 0)
            X "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = ?)", Array("i:" & id)
            X "DELETE FROM TBL_PLAYLIST WHERE ID_USER = ?", Array("i:" & id)
            X "DELETE FROM TBL_USER WHERE ID_USER = ?", Array("i:" & id)
            SetFlash "Usuario eliminado. Tambien se borraron sus " & nPlaylists & " playlist(s).", "ok"
        End If
    End If
    GoToUrl AppRoot() & "admin/users.asp"
End If

If Len(qbuscar) > 0 Then
    Set rs = Q("SELECT U.ID_USER, U.USERNAME, U.EMAIL, U.BIRTHDAY, U.USER_PHOTO, U.USER_TYPE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYLIST P WHERE P.ID_USER = U.ID_USER) AS N " & _
               "FROM TBL_USER U WHERE U.USERNAME LIKE ? OR U.EMAIL LIKE ? ORDER BY U.ID_USER ASC", _
               Array("s:%" & qbuscar & "%", "s:%" & qbuscar & "%"))
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME LIKE ? OR EMAIL LIKE ?", _
                     Array("s:%" & qbuscar & "%", "s:%" & qbuscar & "%")), 0)
Else
    Set rs = Q("SELECT U.ID_USER, U.USERNAME, U.EMAIL, U.BIRTHDAY, U.USER_PHOTO, U.USER_TYPE, " & _
               "(SELECT COUNT(*) FROM TBL_PLAYLIST P WHERE P.ID_USER = U.ID_USER) AS N " & _
               "FROM TBL_USER U ORDER BY U.ID_USER ASC", Empty)
    total = ToInt(XS("SELECT COUNT(*) FROM TBL_USER", Empty), 0)
End If

PageTitle  = "Gestion de usuarios"
PageId     = "admin_users"
PageKicker = "Administracion"
PageSub    = "<b>" & total & "</b> usuario(s) registrado(s)."
%>
<!--#include file="../includes/header.asp"-->

<div class="section">
  <div class="section-head" style="gap:10px;align-items:center;flex-wrap:wrap">
    <form method="get" action="<%=H(AppRoot())%>admin/users.asp" style="display:flex;gap:10px;align-items:center">
      <input class="input" type="text" name="q" value="<%=H(qbuscar)%>" placeholder="Buscar por usuario o correo" style="max-width:300px">
      <button type="submit" class="btn btn-outline btn-sm"><%=Ico("search", 15)%> Buscar</button>
      <% If Len(qbuscar) > 0 Then %><a class="btn btn-outline btn-sm" href="<%=H(AppRoot())%>admin/users.asp">Limpiar</a><% End If %>
    </form>
    <a class="btn btn-primary btn-sm" href="<%=H(AppRoot())%>admin/user_form.asp" style="margin-left:auto"><%=Ico("plus", 15)%> Nuevo usuario</a>
  </div>
</div>

<% If rs.EOF Then %>
  <div class="empty"><span class="ico"><%=Ico("users", 40)%></span><h3>Sin usuarios</h3></div>
<% Else %>
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr>
          <th>ID</th><th>Usuario</th><th>Correo</th><th>Nacimiento</th><th>Tipo</th><th>Playlists</th><th class="actions">Acciones</th>
        </tr>
      </thead>
      <tbody>
      <%
      Do While Not rs.EOF
      %>
      <tr>
        <td><%=rs("ID_USER")%></td>
        <td>
          <div style="display:flex;align-items:center;gap:9px">
            <img class="thumb round" src="<%=H(Resolve(rs("USER_PHOTO"), RES_USERS & "/" & IMG_DEFAULT_USER))%>" alt="">
            <b><%=H(rs("USERNAME"))%></b>
            <% If ToInt(rs("ID_USER"), 0) = CurrentUserID() Then %><span class="tag tag-blue">Tu</span><% End If %>
          </div>
        </td>
        <td><%=H(rs("EMAIL"))%></td>
        <td><%=FmtDate(rs("BIRTHDAY"))%></td>
        <td>
          <% If ToInt(rs("USER_TYPE"), 0) = USER_TYPE_ADMIN Then %>
            <span class="tag tag-green">Administrador</span>
          <% Else %>
            <span class="tag tag-gray">Usuario</span>
          <% End If %>
        </td>
        <td><%=ToInt(rs("N"), 0)%></td>
        <td class="actions">
          <a class="row-btn" title="Editar" href="<%=H(AppRoot())%>admin/user_form.asp?id=<%=rs("ID_USER")%>"><%=Ico("edit", 12)%></a>
          <% If ToInt(rs("ID_USER"), 0) <> CurrentUserID() Then %>
          <form method="post" action="<%=H(AppRoot())%>admin/users.asp" style="display:inline"
                data-confirm="Se eliminara el usuario <%=H(rs("USERNAME"))%> y todas sus playlists. Continuar?">
            <input type="hidden" name="action" value="delete">
            <input type="hidden" name="id" value="<%=rs("ID_USER")%>">
            <button type="submit" class="row-btn danger" title="Eliminar"><%=Ico("trash", 12)%></button>
          </form>
          <% End If %>
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
