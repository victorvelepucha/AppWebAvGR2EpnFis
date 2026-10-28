<%
' ============================================================================
'  PoliMusic - Cabecera comun (barra superior, menu lateral, apertura de main)
'  Variables opcionales que la pagina puede definir antes de incluir este archivo:
'     PageTitle, PageId, PageKicker, PageSub, NoPlayer
' ============================================================================
Dim hdrTitle, hdrPid, hdrKicker, hdrSub, hdrNoPlayer
Dim hdrUsers, hdrSongs, hdrArtists, hdrAlbums, hdrPlaylists
Dim hdrMyLists, hdrUnusedRs, hdrBase, hdrName, hdrPhoto

hdrBase = AppRoot()

If IsEmpty(PageTitle) Or Len(PageTitle & "") = 0 Then
    hdrTitle = APP_NAME
Else
    hdrTitle = PageTitle & " | " & APP_NAME
End If
hdrPid      = CStr(PageId & "")
hdrKicker   = CStr(PageKicker & "")
hdrSub      = CStr(PageSub & "")
hdrNoPlayer = (NoPlayer = True)

hdrUsers     = ToInt(XS("SELECT COUNT(*) FROM TBL_USER", Empty), 0)
hdrSongs     = ToInt(XS("SELECT COUNT(*) FROM TBL_SONG", Empty), 0)
hdrArtists   = ToInt(XS("SELECT COUNT(*) FROM TBL_ARTIST", Empty), 0)
hdrAlbums    = ToInt(XS("SELECT COUNT(*) FROM TBL_ALBUM", Empty), 0)
hdrPlaylists = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty), 0)
hdrName      = CurrentUserName()
hdrPhoto     = CurrentUserPhoto()
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%=H(hdrTitle)%></title>
<link rel="stylesheet" href="<%=H(hdrBase)%>assets/css/style.css">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='16' fill='%231ed760'/%3E%3Cpath d='M12 22V11l9-1.5v11' stroke='%23000' stroke-width='2.4' fill='none' stroke-linecap='round'/%3E%3Ccircle cx='10' cy='22' r='2.6' fill='%23000'/%3E%3Ccircle cx='19' cy='20.5' r='2.6' fill='%23000'/%3E%3C/svg%3E">
</head>
<body>
<div class="app">

  <!-- ======================= Marca ======================= -->
  <div class="brand">
    <div class="brand-logo">PM</div>
    <div class="brand-name">Poli<span>Music</span></div>
  </div>

  <!-- ======================= Barra superior ======================= -->
  <header class="topbar">
    <form class="search-form" method="get" action="<%=H(hdrBase)%>search.asp">
      <span class="ico"><%=Ico("search", 16)%></span>
      <input type="text" name="q" placeholder="Buscar canciones, artistas, albumes o playlists"
             value="<%=H(Request.QueryString("q"))%>" autocomplete="off">
    </form>
    <div class="topbar-spacer"></div>
    <% If IsAdmin() Then %>
      <a class="btn btn-outline btn-sm" href="<%=H(hdrBase)%>admin/default.asp"><%=Ico("gear", 15)%> Administracion</a>
    <% End If %>
    <a href="<%=H(hdrBase)%>profile.asp" class="user-chip" title="Mi perfil">
      <img src="<%=H(hdrPhoto)%>" alt="Foto de perfil">
      <span><%=H(hdrName)%></span>
      <% If IsAdmin() Then %><span class="badge-admin">Admin</span><% End If %>
    </a>
    <a class="btn-logout" href="<%=H(hdrBase)%>logout.asp" title="Cerrar sesion"><%=Ico("logout", 15)%> Salir</a>
  </header>

  <!-- ======================= Menu lateral ======================= -->
  <aside class="sidebar">
    <div class="nav-group">
      <div class="nav-title">Navegar</div>
      <a class="nav-item<%=ActiveIf(hdrPid, "home")%>" href="<%=H(hdrBase)%>home.asp">
        <span class="ico"><%=Ico("home", 18)%></span> Inicio</a>
      <a class="nav-item<%=ActiveIf(hdrPid, "search")%>" href="<%=H(hdrBase)%>search.asp">
        <span class="ico"><%=Ico("search", 18)%></span> Buscar</a>
    </div>

    <div class="nav-group">
      <div class="nav-title">Biblioteca</div>
      <a class="nav-item<%=ActiveIf(hdrPid, "songs")%>" href="<%=H(hdrBase)%>songs.asp">
        <span class="ico"><%=Ico("music", 18)%></span> Canciones <span class="count"><%=hdrSongs%></span></a>
      <a class="nav-item<%=ActiveIf(hdrPid, "artists")%>" href="<%=H(hdrBase)%>artists.asp">
        <span class="ico"><%=Ico("mic", 18)%></span> Artistas <span class="count"><%=hdrArtists%></span></a>
      <a class="nav-item<%=ActiveIf(hdrPid, "albums")%>" href="<%=H(hdrBase)%>albums.asp">
        <span class="ico"><%=Ico("disc", 18)%></span> Albumes <span class="count"><%=hdrAlbums%></span></a>
      <a class="nav-item<%=ActiveIf(hdrPid, "playlists")%>" href="<%=H(hdrBase)%>playlists.asp">
        <span class="ico"><%=Ico("list", 18)%></span> Playlists <span class="count"><%=hdrPlaylists%></span></a>
    </div>

    <div class="nav-sep"></div>

    <div class="nav-group">
      <div class="nav-title">Mis playlists</div>
      <a class="nav-item" href="<%=H(hdrBase)%>playlist_edit.asp">
        <span class="ico"><%=Ico("plus", 18)%></span> Crear playlist</a>
      <%
      Set hdrMyLists = Q("SELECT TOP 8 ID_PLAYLIST, TITLE, TYPE FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY CREATION_DATE DESC, ID_PLAYLIST DESC", _
                        Array("i:" & CurrentUserID()))
      Do While Not hdrMyLists.EOF
      %>
        <a class="side-playlist" href="<%=H(hdrBase)%>playlist.asp?id=<%=hdrMyLists("ID_PLAYLIST")%>">
          <span class="dot"><%=Ico("music", 12)%></span><%=H(Shorten(hdrMyLists("TITLE"), 22))%>
          <% If ToInt(hdrMyLists("TYPE"), 0) = PLAYLIST_PRIVATE Then %><span class="tag tag-gray" style="float:right">Priv</span><% End If %>
        </a>
      <%
        hdrMyLists.MoveNext
      Loop
      hdrMyLists.Close
      If IsAdmin() Then
      %>
        <div class="nav-sep"></div>
        <div class="nav-title">Administracion</div>
        <a class="nav-item<%=ActiveIf(hdrPid, "admin")%>" href="<%=H(hdrBase)%>admin/default.asp">
          <span class="ico"><%=Ico("gear", 18)%></span> Panel general</a>
        <a class="nav-item<%=ActiveIf(hdrPid, "admin_users")%>" href="<%=H(hdrBase)%>admin/users.asp">
          <span class="ico"><%=Ico("users", 18)%></span> Usuarios <span class="count"><%=hdrUsers%></span></a>
        <a class="nav-item<%=ActiveIf(hdrPid, "admin_songs")%>" href="<%=H(hdrBase)%>admin/songs.asp">
          <span class="ico"><%=Ico("music", 18)%></span> Canciones</a>
        <a class="nav-item<%=ActiveIf(hdrPid, "admin_artists")%>" href="<%=H(hdrBase)%>admin/artists.asp">
          <span class="ico"><%=Ico("mic", 18)%></span> Artistas</a>
        <a class="nav-item<%=ActiveIf(hdrPid, "admin_albums")%>" href="<%=H(hdrBase)%>admin/albums.asp">
          <span class="ico"><%=Ico("disc", 18)%></span> Albumes</a>
        <a class="nav-item<%=ActiveIf(hdrPid, "admin_playlists")%>" href="<%=H(hdrBase)%>admin/playlists.asp">
          <span class="ico"><%=Ico("list", 18)%></span> Playlists</a>
      <%
      End If
      %>
    </div>
  </aside>

  <!-- ======================= Contenido ======================= -->
  <main class="main">
    <div class="main-inner">
      <%=TakeFlash()%>
      <%=QsAlert()%>
      <% If Len(hdrKicker) > 0 Or Len(hdrSub) > 0 Then %>
        <div class="page-head">
          <div class="head-text">
            <% If Len(hdrKicker) > 0 Then %><div class="kicker"><%=H(hdrKicker)%></div><% End If %>
            <h1><%=H(CStr(PageTitle & ""))%></h1>
            <% If Len(hdrSub) > 0 Then %><div class="sub"><%=hdrSub%></div><% End If %>
          </div>
        </div>
      <% End If %>
