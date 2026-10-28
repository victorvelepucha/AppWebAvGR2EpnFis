<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim errCode, errDesc, errUrl
errCode = CStr(Request.QueryString("code") & "")
errDesc = CStr(Request.QueryString("msg") & "")
errUrl  = CStr(Request.QueryString("url") & "")
If Len(errDesc) = 0 Then errDesc = "No tiene permisos para ver esta pagina o el recurso no existe."
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Acceso denegado | <%=H(APP_NAME)%></title>
<link rel="stylesheet" href="<%=H(AppRoot())%>assets/css/style.css">
</head>
<body>
<div class="auth-wrap">
  <div class="auth-card" style="max-width:520px;text-align:center">
    <div class="auth-brand">
      <div class="brand-logo">PM</div>
      <div class="brand-name">Poli<span>Music</span></div>
    </div>
    <h2 style="margin-top:18px">
      <% If Len(errCode) > 0 Then %><%=H(errCode)%> - <% End If %>Acceso restringido
    </h2>
    <p style="color:#a7a7a7"><%=H(errDesc)%></p>
    <% If Len(errUrl) > 0 Then %>
      <p style="color:#6f6f6f;font-size:12px;word-break:break-all"><%=H(errUrl)%></p>
    <% End If %>
    <div style="display:flex;gap:10px;justify-content:center;margin-top:22px">
      <a class="btn btn-outline" href="<%=H(AppRoot())%>home.asp"><%=Ico("back", 16)%> Volver al inicio</a>
      <a class="btn btn-primary" href="<%=H(AppRoot())%>logout.asp"><%=Ico("logout", 16)%> Cambiar de usuario</a>
    </div>
  </div>
</div>
</body>
</html>
