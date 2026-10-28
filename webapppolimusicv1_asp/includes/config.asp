<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================================================
'  PoliMusic - Configuracion general de la aplicacion
'  Proyecto: Clon de Spotify en ASP clasico + SQL Server
' ============================================================================
Option Explicit

Response.CodePage = 65001
Response.Charset  = "utf-8"
Session.CodePage  = 65001

' --- Identidad de la aplicacion -------------------------------------------
Const APP_NAME     = "PoliMusic"
Const APP_TAGLINE  = "Tu musica, tus playlists, tus artistas"

' --- Conexion a SQL Server -------------------------------------------------
Const DB_SERVER    = "localhost"
Const DB_NAME      = "BDD_PoliMusic"
Const DB_USER      = "usr_polimusic"
Const DB_PASS      = "Politecnica1"

' --- Recursos multimedia ---------------------------------------------------
' Carpeta (dentro del sitio) donde viven los mp3 y las imagenes.
Const RES_FOLDER   = "SongFilesAndImages_resources"

' Subcarpetas reales de recursos
Const RES_SONGS    = "SongFiles"
Const RES_USERS    = "UserImages"
Const RES_PLAYLISTS= "PlaylistImages"
Const RES_ALBUMS   = "AlbumImages"
Const RES_ARTISTS  = "ArtistImages"

' Imagenes por defecto cuando el recurso referenciado no existe
Const IMG_DEFAULT_USER     = "default.jpg"
Const IMG_DEFAULT_PLAYLIST = "default.jpg"

' --- Catalogos (TBL_CATALOG) ----------------------------------------------
' COD_CATALOG_PARENT = '1'  ->  Tipos de usuario
Const USER_TYPE_NORMAL = 0
Const USER_TYPE_ADMIN  = 1

' COD_CATALOG_PARENT = '2'  ->  Tipos de playlist
Const PLAYLIST_PUBLIC  = 0
Const PLAYLIST_PRIVATE = 1

' --- Varios ----------------------------------------------------------------
Const ROWS_PER_PAGE = 12
Const MAX_LOGIN_ATTEMPTS = 5

' --- Variables de presentacion --------------------------------------------
' Cada pagina puede asignarlas ANTES de incluir header.asp / footer.asp.
Dim PageTitle    ' Titulo del documento y de la pestana
Dim PageId       ' Identificador para resaltar la opcion activa del menu
Dim PageKicker   ' Texto pequeno sobre el titulo
Dim PageSub      ' Subtitulo / descripcion
Dim NoPlayer     ' True para no mostrar la barra del reproductor
%>
