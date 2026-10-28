<%
' ============================================================================
'  PoliMusic - Libreria de acceso a datos y utilidades
' ============================================================================
' NOTA: requiere que config.asp se haya incluido antes.

Dim gConn, gAppRoot, gAppPhys

' ---------------------------------------------------------------------------
'  Rutas de la aplicacion
' ---------------------------------------------------------------------------

' Devuelve la raiz virtual de la app, siempre terminada en "/".
'   /polimusicweb/login.asp        -> /polimusicweb/
'   /polimusicweb/admin/users.asp  -> /polimusicweb/
Function AppRoot()
    Dim p, i
    If Len(gAppRoot & "") > 0 Then
        AppRoot = gAppRoot
        Exit Function
    End If
    p = Request.ServerVariables("SCRIPT_NAME")
    i = InStrRev(p, "/")
    If i > 0 Then
        p = Left(p, i)
    Else
        p = "/"
    End If
    If LCase(Right(p, 7)) = "/admin/" Then p = Left(p, Len(p) - 6)
    gAppRoot = p
    AppRoot = gAppRoot
End Function

' Devuelve la raiz fisica de la app, siempre terminada en "\".
Function AppPhys()
    Dim root
    If Len(gAppPhys & "") > 0 Then
        AppPhys = gAppPhys
        Exit Function
    End If
    root = Server.MapPath("/")
    If Right(root, 1) <> "\" Then root = root & "\"
    gAppPhys = root & Replace(Mid(AppRoot(), 2), "/", "\")
    AppPhys = gAppPhys
End Function

' Ruta fisica de un archivo de recurso (relativa a la carpeta de recursos).
Function ResPhys(rel)
    ResPhys = AppPhys() & RES_FOLDER & "\" & Replace(rel, "/", "\")
End Function

' URL de un archivo de recurso (relativa a la carpeta de recursos).
Function ResUrl(rel)
    ResUrl = AppRoot() & RES_FOLDER & "/" & rel
End Function

' ---------------------------------------------------------------------------
'  Conexion a la base de datos
' ---------------------------------------------------------------------------

' Lista de cadenas de conexion, en orden de preferencia.
Function ConnStrList()
    Dim a(4)
    a(0) = "Provider=MSOLEDBSQL;Server=" & DB_SERVER & ";Database=" & DB_NAME & _
           ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";Encrypt=Optional;TrustServerCertificate=Yes;"
    a(1) = "Provider=MSOLEDBSQL19;Server=" & DB_SERVER & ";Database=" & DB_NAME & _
           ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";Encrypt=Optional;TrustServerCertificate=Yes;"
    a(2) = "Provider=SQLOLEDB;Data Source=" & DB_SERVER & ";Initial Catalog=" & DB_NAME & _
           ";User ID=" & DB_USER & ";Password=" & DB_PASS & ";"
    a(3) = "Driver={ODBC Driver 18 for SQL Server};Server=" & DB_SERVER & ";Database=" & DB_NAME & _
           ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & ";Encrypt=no;TrustServerCertificate=yes;"
    a(4) = "Driver={ODBC Driver 17 for SQL Server};Server=" & DB_SERVER & ";Database=" & DB_NAME & _
           ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & ";Encrypt=no;TrustServerCertificate=yes;"
    ConnStrList = a
End Function

' Abre (una sola vez por peticion) la conexion a la base de datos.
Function Db()
    Dim list, i, c
    If Not IsEmpty(gConn) Then
        Set Db = gConn
        Exit Function
    End If
    list = ConnStrList()
    For i = 0 To UBound(list)
        On Error Resume Next
        Set c = Server.CreateObject("ADODB.Connection")
        c.ConnectionTimeout = 10
        c.CommandTimeout = 30
        c.Open list(i)
        If Err.Number = 0 Then
            On Error GoTo 0
            Set gConn = c
            Set Db = gConn
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next
    Response.Write "<h2 style=""font-family:sans-serif"">Error de conexion</h2>" & _
                   "<p style=""font-family:sans-serif"">No fue posible conectar con SQL Server " & _
                   "(" & DB_SERVER & " / " & DB_NAME & "). Verifique que el servicio este activo " & _
                   "y que el login '" & DB_USER & "' exista.</p>"
    Response.End
End Function

' ---------------------------------------------------------------------------
'  Ejecucion de consultas parametrizadas
'  Convencion de parametros: Array("i:12", "s:texto", "d:2026-01-31", "n:")
'     i: -> entero       s: -> cadena       d: -> fecha       n: -> NULL
' ---------------------------------------------------------------------------
Function MakeCmd(sql, params)
    Dim c, i, v, t
    Set c = Server.CreateObject("ADODB.Command")
    Set c.ActiveConnection = Db()
    c.CommandType = 1        ' adCmdText
    c.CommandText = sql
    If IsArray(params) Then
        For i = 0 To UBound(params)
            v = params(i)
            If VarType(v) = vbNullString Or Len(CStr(v)) < 2 Then
                t = "s"
                v = ""
            Else
                t = LCase(Left(v, 1))
                v = Mid(v, 3)
            End If
            Select Case t
                Case "i"
                    If Len(v) = 0 Then
                        c.Parameters.Append c.CreateParameter("p" & i, 3, 1, , Null)
                    Else
                        c.Parameters.Append c.CreateParameter("p" & i, 3, 1, , CLng(v))
                    End If
                Case "d"
                    c.Parameters.Append c.CreateParameter("p" & i, 135, 1, , CDate(v))
                Case "n"
                    c.Parameters.Append c.CreateParameter("p" & i, 202, 1, 4000, Null)
                Case Else
                    c.Parameters.Append c.CreateParameter("p" & i, 202, 1, 4000, CStr(v))
            End Select
        Next
    End If
    Set MakeCmd = c
End Function

' Devuelve un Recordset desconectado.
Function Q(sql, params)
    Dim c, rs
    Set c = MakeCmd(sql, params)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3      ' adUseClient
    rs.Open c
    Set rs.ActiveConnection = Nothing
    Set Q = rs
End Function

' Ejecuta una sentencia y devuelve el numero de filas afectadas.
Function X(sql, params)
    Dim c, n
    Set c = MakeCmd(sql, params)
    c.Execute n
    X = n
End Function

' Ejecuta una consulta y devuelve el primer campo de la primera fila.
Function XS(sql, params)
    Dim rs
    Set rs = Q(sql, params)
    If rs.EOF Then
        XS = Null
    Else
        If IsNull(rs(0)) Then XS = Null Else XS = rs(0)
    End If
    rs.Close
End Function

' Devuelve el ultimo IDENTITY insertado en la conexion actual.
Function LastId()
    LastId = CLng(XS("SELECT CAST(SCOPE_IDENTITY() AS INT)", Empty))
End Function

' ---------------------------------------------------------------------------
'  Utilidades generales
' ---------------------------------------------------------------------------

' Escapa texto para mostrarlo en HTML.
Function H(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        H = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

' Escapa texto para usarlo dentro de un atributo JavaScript.
Function Js(v)
    Dim s
    If IsNull(v) Then
        Js = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "<", "\u003c")
    s = Replace(s, ">", "\u003e")
    Js = s
End Function

' Convierte a entero de forma segura.
Function ToInt(v, def)
    If IsNull(v) Or IsEmpty(v) Then
        ToInt = def
        Exit Function
    End If
    If Not IsNumeric(v) Then
        ToInt = def
        Exit Function
    End If
    On Error Resume Next
    ToInt = CLng(v)
    If Err.Number <> 0 Then
        Err.Clear
        ToInt = def
    End If
    On Error GoTo 0
End Function

' Recorta un texto y agrega puntos suspensivos.
Function Shorten(v, n)
    Dim s
    s = Trim(CStr(v & ""))
    If Len(s) > n Then s = Left(s, n - 1) & "..."
    Shorten = s
End Function

' Formatea una fecha para mostrarla.
Function FmtDate(v)
    If IsNull(v) Or IsEmpty(v) Then
        FmtDate = "-"
        Exit Function
    End If
    On Error Resume Next
    FmtDate = Right("0" & Day(v), 2) & "/" & Right("0" & Month(v), 2) & "/" & Year(v)
    If Err.Number <> 0 Then
        Err.Clear
        FmtDate = "-"
    End If
    On Error GoTo 0
End Function

' Formatea una fecha para un input type="date".
Function FmtDateInput(v)
    If IsNull(v) Or IsEmpty(v) Then
        FmtDateInput = ""
        Exit Function
    End If
    On Error Resume Next
    FmtDateInput = Year(v) & "-" & Right("0" & Month(v), 2) & "-" & Right("0" & Day(v), 2)
    If Err.Number <> 0 Then
        Err.Clear
        FmtDateInput = ""
    End If
    On Error GoTo 0
End Function

' Devuelve " active" cuando el identificador de pagina coincide (menu lateral).
Function ActiveIf(pageId, target)
    If LCase(Trim(CStr(pageId & ""))) = LCase(Trim(CStr(target & ""))) Then
        ActiveIf = " active"
    Else
        ActiveIf = ""
    End If
End Function

' Redirige y detiene la ejecucion.
Sub GoToUrl(url)
    Response.Redirect url
    Response.End
End Sub

' ---------------------------------------------------------------------------
'  Recursos multimedia: resolucion de rutas
' ---------------------------------------------------------------------------

' Traduce el nombre logico de carpeta usado en la BD a la carpeta real.
Function MapFolder(name)
    Dim n
    n = LCase(Trim(CStr(name & "")))
    Select Case n
        Case "songfiles", "songfile", "songs"       : MapFolder = RES_SONGS
        Case "userimages", "userimage", "users"      : MapFolder = RES_USERS
        Case "playlistimages", "playlistimage"       : MapFolder = RES_PLAYLISTS
        Case "albumimages", "albumimage", "albums"   : MapFolder = RES_ALBUMS
        Case "artistimages", "artistimage", "artists": MapFolder = RES_ARTISTS
        Case Else                                    : MapFolder = ""
    End Select
End Function

' Resuelve la ruta almacenada en la BD a una URL utilizable.
' Devuelve "" cuando el archivo no existe y no hay imagen por defecto.
' defaultRel: ruta relativa (dentro de la carpeta de recursos) a usar como respaldo.
Function Resolve(dbPath, defaultRel)
    Dim p, i, folder, file, realFolder, rel, fso
    Resolve = ""
    p = Trim(CStr(dbPath & ""))
    p = Replace(p, "\", "/")
    Do While Left(p, 1) = "~" Or Left(p, 1) = "/"
        p = Mid(p, 2)
    Loop
    If Len(p) = 0 Then
        If Len(defaultRel) > 0 Then Resolve = ResUrl(defaultRel)
        Exit Function
    End If
    i = InStrRev(p, "/")
    If i > 0 Then
        folder = Left(p, i - 1)
        file = Mid(p, i + 1)
    Else
        folder = ""
        file = p
    End If
    realFolder = MapFolder(folder)
    If Len(realFolder) = 0 Then realFolder = folder
    If Len(file) > 0 And Len(realFolder) > 0 Then
        rel = realFolder & "/" & file
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(ResPhys(rel)) Then
            Resolve = ResUrl(rel)
            Set fso = Nothing
            Exit Function
        End If
        Set fso = Nothing
    End If
    If Len(defaultRel) > 0 Then Resolve = ResUrl(defaultRel)
End Function

' Indica si el archivo de audio referenciado por una cancion existe fisicamente.
Function SongExists(songPath)
    Dim p, i, folder, file, rel, fso
    SongExists = False
    p = Replace(Trim(CStr(songPath & "")), "\", "/")
    Do While Left(p, 1) = "~" Or Left(p, 1) = "/"
        p = Mid(p, 2)
    Loop
    i = InStrRev(p, "/")
    If i > 0 Then
        folder = Left(p, i - 1)
        file = Mid(p, i + 1)
    Else
        folder = "songfiles"
        file = p
    End If
    If Len(file) = 0 Then Exit Function
    If Len(MapFolder(folder)) > 0 Then folder = MapFolder(folder)
    rel = folder & "/" & file
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    SongExists = fso.FileExists(ResPhys(rel))
    Set fso = Nothing
End Function

' Lista los archivos de una carpeta de recursos (para los selectores del admin).
Function ListResFiles(folder)
    Dim fso, f, out, arr(), n, i, j, tmp, fld
    fld = MapFolder(folder)
    If Len(fld) = 0 Then fld = folder
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    n = 0
    If fso.FolderExists(ResPhys(fld)) Then
        Set f = fso.GetFolder(ResPhys(fld))
        For Each tmp In f.Files
            ReDim Preserve arr(n)
            arr(n) = tmp.Name
            n = n + 1
        Next
        Set f = Nothing
    End If
    Set fso = Nothing
    ' ordenamiento burbuja simple
    If n > 1 Then
        For i = 0 To n - 2
            For j = 0 To n - 2 - i
                If LCase(arr(j)) > LCase(arr(j + 1)) Then
                    tmp = arr(j)
                    arr(j) = arr(j + 1)
                    arr(j + 1) = tmp
                End If
            Next
        Next
    End If
    If n = 0 Then
        ListResFiles = Array()
    Else
        ListResFiles = arr
    End If
End Function

' ---------------------------------------------------------------------------
'  Iconos SVG en linea
' ---------------------------------------------------------------------------
Function Ico(name, size)
    Dim p, s
    If IsEmpty(size) Or IsNull(size) Then size = 18
    s = CStr(size)
    Select Case LCase(Trim(CStr(name & "")))
        Case "home"    : p = "<path d=""M3 11l9-8 9 8v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z""/>"
        Case "music"   : p = "<path d=""M9 18V5l12-2v13""/><circle cx=""6"" cy=""18"" r=""3""/><circle cx=""18"" cy=""16"" r=""3""/>"
        Case "mic"     : p = "<path d=""M12 15a4 4 0 0 0 4-4V6a4 4 0 0 0-8 0v5a4 4 0 0 0 4 4z""/><path d=""M19 11a7 7 0 0 1-14 0""/><line x1=""12"" y1=""18"" x2=""12"" y2=""22""/>"
        Case "disc"    : p = "<circle cx=""12"" cy=""12"" r=""9""/><circle cx=""12"" cy=""12"" r=""3""/>"
        Case "list"    : p = "<line x1=""8"" y1=""6"" x2=""21"" y2=""6""/><line x1=""8"" y1=""12"" x2=""21"" y2=""12""/><line x1=""8"" y1=""18"" x2=""21"" y2=""18""/><circle cx=""4"" cy=""6"" r=""1""/><circle cx=""4"" cy=""12"" r=""1""/><circle cx=""4"" cy=""18"" r=""1""/>"
        Case "gear"    : p = "<circle cx=""12"" cy=""12"" r=""3""/><path d=""M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.9.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.9 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.9l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.9.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.9-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.9V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z""/>"
        Case "search"  : p = "<circle cx=""11"" cy=""11"" r=""7""/><line x1=""20"" y1=""20"" x2=""16.5"" y2=""16.5""/>"
        Case "play"    : p = "<polygon points=""6,4 20,12 6,20"" fill=""currentColor"" stroke=""none""/>"
        Case "pause"   : p = "<rect x=""6"" y=""4"" width=""4"" height=""16"" fill=""currentColor"" stroke=""none""/><rect x=""14"" y=""4"" width=""4"" height=""16"" fill=""currentColor"" stroke=""none""/>"
        Case "next"    : p = "<polygon points=""5,4 15,12 5,20"" fill=""currentColor"" stroke=""none""/><rect x=""17"" y=""4"" width=""2.5"" height=""16"" fill=""currentColor"" stroke=""none""/>"
        Case "prev"    : p = "<polygon points=""19,4 9,12 19,20"" fill=""currentColor"" stroke=""none""/><rect x=""4.5"" y=""4"" width=""2.5"" height=""16"" fill=""currentColor"" stroke=""none""/>"
        Case "shuffle" : p = "<polyline points=""16,3 21,3 21,8""/><line x1=""4"" y1=""20"" x2=""21"" y2=""3""/><polyline points=""21,16 21,21 16,21""/><line x1=""15"" y1=""15"" x2=""21"" y2=""21""/><line x1=""4"" y1=""4"" x2=""9"" y2=""9""/>"
        Case "repeat"  : p = "<polyline points=""17,1 21,5 17,9""/><path d=""M3 11V9a4 4 0 0 1 4-4h14""/><polyline points=""7,23 3,19 7,15""/><path d=""M21 13v2a4 4 0 0 1-4 4H3""/>"
        Case "volume"  : p = "<polygon points=""11,5 6,9 2,9 2,15 6,15 11,19""/><path d=""M15.5 8.5a5 5 0 0 1 0 7""/><path d=""M18.5 5.5a9 9 0 0 1 0 13""/>"
        Case "user"    : p = "<circle cx=""12"" cy=""8"" r=""4""/><path d=""M4 21a8 8 0 0 1 16 0""/>"
        Case "users"   : p = "<path d=""M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2""/><circle cx=""9"" cy=""7"" r=""4""/><path d=""M23 21v-2a4 4 0 0 0-3-3.87""/><path d=""M16 3.13a4 4 0 0 1 0 7.75""/>"
        Case "logout"  : p = "<path d=""M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4""/><polyline points=""16,17 21,12 16,7""/><line x1=""21"" y1=""12"" x2=""9"" y2=""12""/>"
        Case "plus"    : p = "<line x1=""12"" y1=""5"" x2=""12"" y2=""19""/><line x1=""5"" y1=""12"" x2=""19"" y2=""12""/>"
        Case "edit"    : p = "<path d=""M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7""/><path d=""M18.5 2.5a2.1 2.1 0 0 1 3 3L12 15l-4 1 1-4z""/>"
        Case "trash"   : p = "<polyline points=""3,6 5,6 21,6""/><path d=""M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6""/><path d=""M10 11v6M14 11v6""/>"
        Case "back"    : p = "<line x1=""19"" y1=""12"" x2=""5"" y2=""12""/><polyline points=""12,19 5,12 12,5""/>"
        Case "clock"   : p = "<circle cx=""12"" cy=""12"" r=""9""/><polyline points=""12,7 12,12 15.5,14""/>"
        Case "chart"   : p = "<line x1=""18"" y1=""20"" x2=""18"" y2=""10""/><line x1=""12"" y1=""20"" x2=""12"" y2=""4""/><line x1=""6"" y1=""20"" x2=""6"" y2=""14""/>"
        Case "check"   : p = "<polyline points=""20,6 9,17 4,12""/>"
        Case "x"       : p = "<line x1=""18"" y1=""6"" x2=""6"" y2=""18""/><line x1=""6"" y1=""6"" x2=""18"" y2=""18""/>"
        Case "save"    : p = "<path d=""M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z""/><polyline points=""17,21 17,13 7,13 7,21""/><polyline points=""7,3 7,8 15,8""/>"
        Case Else      : p = "<circle cx=""12"" cy=""12"" r=""9""/>"
    End Select
    Ico = "<svg class=""ico-svg"" width=""" & s & """ height=""" & s & """ viewBox=""0 0 24 24"" " & _
          "fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" " & _
          "stroke-linejoin=""round"" aria-hidden=""true"">" & p & "</svg>"
End Function

' ---------------------------------------------------------------------------
'  Portadas generadas (cuando no existe una imagen real)
' ---------------------------------------------------------------------------
Function CoverPlaceholder(text, extraClass)
    Dim seed, hues, idx, hue, initials, words, i, w
    seed = 0
    For i = 1 To Len(CStr(text & ""))
        seed = seed + AscW(Mid(CStr(text & ""), i, 1)) * i
    Next
    If seed < 0 Then seed = -seed
    hue = seed Mod 360
    initials = ""
    words = Split(Trim(CStr(text & "")), " ")
    For i = 0 To UBound(words)
        w = Trim(words(i))
        If Len(w) > 0 Then
            initials = initials & UCase(Left(w, 1))
            If Len(initials) >= 2 Then Exit For
        End If
    Next
    If Len(initials) = 0 Then initials = "?"
    CoverPlaceholder = "<div class=""cover-ph " & CStr(extraClass & "") & """ style=""background:linear-gradient(135deg," & _
        "hsl(" & hue & ",62%,42%),hsl(" & ((hue + 48) Mod 360) & ",58%,24%))"">" & H(initials) & "</div>"
End Function

' ---------------------------------------------------------------------------
'  Fragmentos SQL reutilizables
' ---------------------------------------------------------------------------

' SELECT base de canciones con la lista de artistas concatenada.
Function SqlSongs()
    SqlSongs = "SELECT S.ID_SONG, S.SONG_NAME, S.SONG_PATH, S.PLAYS, " & _
        "ISNULL(STUFF((SELECT ', ' + A.ARTIST_NAME + ISNULL(' ' + A.ARTIST_LASTNAME, '') " & _
        "FROM TBL_PLAYER P INNER JOIN TBL_ARTIST A ON A.ID_ARTIST = P.ID_ARTIST " & _
        "WHERE P.ID_SONG = S.ID_SONG FOR XML PATH('')), 1, 2, ''), 'Artista desconocido') AS ARTISTAS " & _
        "FROM TBL_SONG S "
End Function

' ---------------------------------------------------------------------------
'  Componentes de interfaz
' ---------------------------------------------------------------------------

' Dibuja una fila de cancion dentro de una tabla .tracklist.
'   pi        : indice de reproduccion (0..n) o -1 si la cancion no tiene archivo
'   songId    : ID de la cancion
'   sName     : nombre de la cancion
'   sArtist   : artistas concatenados
'   plays     : numero de reproducciones
'   url       : URL del mp3 ("" si no existe)
'   cover     : URL de la portada ("" para generar una)
'   actions   : HTML adicional de acciones (editar, quitar de playlist, ...)
Function SongRow(pi, songId, sName, sArtist, plays, url, cover, actions)
    Dim html, cov, dataAttrs, playBtn, cls
    cls = "js-track"
    If Len(url) = 0 Then cls = cls & " no-audio"

    If Len(cover) > 0 Then
        cov = "<img src=""" & H(cover) & """ alt="""" loading=""lazy"">"
    Else
        cov = CoverPlaceholder(sName, "sm")
    End If

    dataAttrs = ""
    If Len(url) > 0 Then
        dataAttrs = " data-song-url=""" & H(url) & """" & _
                    " data-song-id=""" & H(songId) & """" & _
                    " data-song-title=""" & H(sName) & """" & _
                    " data-song-artist=""" & H(sArtist) & """" & _
                    " data-song-cover=""" & H(cover) & """"
        playBtn = "<button type=""button"" class=""row-btn play"" data-play-index=""" & pi & """ title=""Reproducir"">" & Ico("play", 12) & "</button>"
    Else
        playBtn = "<button type=""button"" class=""row-btn"" disabled title=""Archivo MP3 no disponible"">" & Ico("play", 12) & "</button>"
    End If

    html = "<tr class=""" & cls & """ data-row-song=""" & H(songId) & """" & dataAttrs & ">" & _
           "<td class=""idx"">" & (pi + 1) & "</td>" & _
           "<td class=""t-cover"">" & cov & "</td>" & _
           "<td><div class=""t-name"">" & H(sName) & "</div>" & _
           "<div class=""t-sub"">" & H(sArtist) & "</div></td>" & _
           "<td class=""num"">" & ToInt(plays, 0) & "</td>" & _
           "<td class=""t-actions"">" & playBtn & actions & "</td></tr>"
    SongRow = html
End Function

' ---------------------------------------------------------------------------
'  Contrasenas
' ---------------------------------------------------------------------------

' Calcula el hash SHA1 (hexadecimal en mayusculas) de una cadena.
' Se apoya en SQL Server (HASHBYTES) para evitar implementar SHA1 en VBScript.
Function HashPassword(plain)
    Dim r
    r = XS("SELECT UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', CONVERT(VARCHAR(200), ?)), 2))", Array("s:" & plain))
    If IsNull(r) Then
        HashPassword = ""
    Else
        HashPassword = UCase(Trim(CStr(r)))
    End If
End Function

' Comprueba una contrasena contra el valor almacenado.
' Acepta tanto valores en texto plano (datos heredados) como hashes SHA1.
Function CheckPassword(plain, stored)
    Dim s
    CheckPassword = False
    If IsNull(stored) Then Exit Function
    s = Trim(CStr(stored))
    If s = CStr(plain) Then
        CheckPassword = True
        Exit Function
    End If
    If Len(s) = 40 Then
        If UCase(s) = HashPassword(plain) Then CheckPassword = True
    End If
End Function
%>
