<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Agrega o quita una cancion de una playlist.
RequireLogin

Dim action, plId, songId, back, rsP, owner, isOwner, exists, msg

action = LCase(Trim(CStr(Request.Form("action") & "")))
plId   = ToInt(Request.Form("playlist"), 0)
songId = ToInt(Request.Form("song"), 0)
back   = Trim(CStr(Request.Form("back") & ""))
msg    = ""

If plId <= 0 Or songId <= 0 Then
    GoToUrl AppRoot() & "playlists.asp?err=" & Server.URLEncode("Solicitud incompleta.")
End If

Set rsP = Q("SELECT ID_PLAYLIST, ID_USER, TITLE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array("i:" & plId))
If rsP.EOF Then
    rsP.Close
    GoToUrl AppRoot() & "playlists.asp?err=" & Server.URLEncode("La playlist indicada no existe.")
End If

isOwner = (ToInt(rsP("ID_USER"), 0) = CurrentUserID())
If Not isOwner And Not IsAdmin() Then
    rsP.Close
    GoToUrl AppRoot() & "error.asp?code=403&msg=" & Server.URLEncode("Solo el propietario puede modificar esta playlist.")
End If

' Verificar que la cancion exista
exists = ToInt(XS("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = ?", Array("i:" & songId)), 0)
If exists = 0 Then
    rsP.Close
    GoToUrl AppRoot() & "playlists.asp?err=" & Server.URLEncode("La cancion indicada no existe.")
End If

If action = "remove" Then
    X "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array("i:" & plId, "i:" & songId)
    msg = "Cancion quitada de la playlist."
Else
    exists = ToInt(XS("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", _
                      Array("i:" & plId, "i:" & songId)), 0)
    If exists = 0 Then
        X "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array("i:" & plId, "i:" & songId)
        msg = "Cancion agregada a """ & CStr(rsP("TITLE") & "") & """."
    Else
        msg = "La cancion ya estaba en la playlist."
    End If
End If

rsP.Close
SetFlash msg, "ok"

If Len(back) > 0 And InStr(back, "//") = 0 And InStr(back, ":") = 0 Then
    GoToUrl AppRoot() & back
Else
    GoToUrl AppRoot() & "playlist.asp?id=" & plId
End If
%>
