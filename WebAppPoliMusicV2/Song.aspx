<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Song list - PoliMusic</title>
    <style type="text/css">
        :root {
            --bg: #0b0f19;
            --bg-soft: #0f1522;
            --surface: #121a2a;
            --line: rgba(255, 255, 255, 0.07);
            --line-strong: rgba(255, 255, 255, 0.12);
            --text: #f2f5fa;
            --text-soft: #c3cbdb;
            --muted: #8b96ad;
            --blue: #2e7cf6;
            --blue-light: #6ea8ff;
            --blue-deep: #1b5fd9;
            color-scheme: dark;
        }

        * {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
        }

        body {
            min-height: 100vh;
            background: radial-gradient(1100px 520px at 12% -12%, rgba(46, 124, 246, 0.22), transparent 60%), radial-gradient(900px 500px at 92% 0%, rgba(27, 95, 217, 0.18), transparent 55%), var(--bg);
            color: var(--text);
            font-family: "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
            -webkit-font-smoothing: antialiased;
        }

        /* ---------- Top bar ---------- */
        .app-bar {
            position: sticky;
            top: 0;
            z-index: 20;
            background: rgba(11, 15, 25, 0.82);
            border-bottom: 1px solid var(--line);
            backdrop-filter: blur(14px);
        }

        .app-bar-inner {
            max-width: 1120px;
            margin: 0 auto;
            padding: 14px 26px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 18px;
            flex-wrap: wrap;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 11px;
            font-size: 18px;
            font-weight: 700;
            letter-spacing: -0.2px;
        }

        .brand-mark {
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 17px;
            color: #ffffff;
            background: linear-gradient(135deg, var(--blue-light), var(--blue-deep));
            box-shadow: 0 8px 22px rgba(46, 124, 246, 0.45);
        }

        .brand-sub {
            color: var(--muted);
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 0.16em;
            text-transform: uppercase;
        }

        .user-area {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .avatar {
            width: 38px;
            height: 38px;
            border-radius: 50%;
            object-fit: cover;
            border: 2px solid rgba(46, 124, 246, 0.55);
            background: var(--surface);
        }

        .welcome {
            font-size: 14px;
            color: var(--muted);
        }

        .user-name {
            color: var(--text);
            font-weight: 700;
            letter-spacing: 0.2px;
        }

        .logout {
            display: inline-flex;
            align-items: center;
            gap: 7px;
            padding: 9px 16px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 600;
            text-decoration: none;
            color: var(--blue-light);
            border: 1px solid rgba(46, 124, 246, 0.45);
            background: rgba(46, 124, 246, 0.10);
            transition: background 0.18s ease, color 0.18s ease, transform 0.18s ease, box-shadow 0.18s ease;
        }

        .logout:hover {
            color: #ffffff;
            background: linear-gradient(135deg, var(--blue), var(--blue-deep));
            box-shadow: 0 10px 24px rgba(46, 124, 246, 0.35);
            transform: translateY(-1px);
        }

        /* ---------- Page ---------- */
        .page {
            max-width: 1120px;
            margin: 0 auto;
            padding: 40px 26px 70px;
        }

        .page-head {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 16px;
            flex-wrap: wrap;
            margin-bottom: 22px;
        }

        .page-head h1 {
            margin: 0;
            font-size: 34px;
            line-height: 1.15;
            letter-spacing: -0.8px;
            font-weight: 800;
        }

        .page-head h1 .accent {
            color: var(--blue-light);
        }

        .subtitle {
            margin: 8px 0 0;
            font-size: 14px;
            color: var(--muted);
        }

        .counter {
            display: inline-flex;
            align-items: baseline;
            gap: 8px;
            padding: 10px 18px;
            border-radius: 999px;
            border: 1px solid var(--line-strong);
            background: rgba(255, 255, 255, 0.03);
            font-size: 13px;
            color: var(--muted);
        }

        .counter b {
            color: var(--blue-light);
            font-size: 17px;
            font-weight: 800;
        }

        /* ---------- Table ---------- */
        .table-card {
            background: linear-gradient(180deg, rgba(18, 26, 42, 0.95), rgba(13, 19, 31, 0.95));
            border: 1px solid var(--line);
            border-radius: 20px;
            overflow: hidden;
            box-shadow: 0 30px 70px rgba(0, 0, 0, 0.45);
        }

        table.songs {
            width: 100%;
            border-collapse: collapse;
            border-spacing: 0;
        }

        table.songs thead th {
            padding: 16px 20px;
            text-align: left;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            color: var(--muted);
            background: rgba(255, 255, 255, 0.025);
            border-bottom: 1px solid var(--line);
            white-space: nowrap;
        }

        table.songs tbody tr {
            border-bottom: 1px solid var(--line);
            transition: background 0.18s ease;
        }

        table.songs tbody tr:last-child {
            border-bottom: none;
        }

        table.songs tbody tr:hover {
            background: rgba(46, 124, 246, 0.09);
        }

        table.songs tbody td {
            padding: 14px 20px;
            font-size: 14px;
            color: var(--text-soft);
            vertical-align: middle;
        }

        .col-id {
            width: 96px;
        }

        .id-pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 30px;
            height: 30px;
            padding: 0 8px;
            border-radius: 9px;
            font-size: 12px;
            font-weight: 800;
            color: var(--blue-light);
            background: rgba(46, 124, 246, 0.14);
            border: 1px solid rgba(46, 124, 246, 0.28);
        }

        .song-name {
            color: var(--text);
            font-weight: 700;
            font-size: 15px;
            letter-spacing: -0.1px;
        }

        .col-path {
            width: 34%;
        }

        .path {
            display: inline-block;
            max-width: 100%;
            padding: 5px 11px;
            border-radius: 8px;
            font-family: Consolas, "Courier New", monospace;
            font-size: 12px;
            color: #9fb0cc;
            background: rgba(255, 255, 255, 0.035);
            border: 1px solid var(--line);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            vertical-align: middle;
        }

        .col-plays {
            width: 130px;
            text-align: center;
        }

        table.songs thead th.col-plays {
            text-align: center;
        }

        .plays {
            display: inline-block;
            min-width: 42px;
            padding: 5px 12px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 700;
            color: #d7e6ff;
            background: rgba(46, 124, 246, 0.12);
            border: 1px solid rgba(46, 124, 246, 0.3);
        }

        .col-play {
            width: 300px;
        }

        .player {
            width: 100%;
            max-width: 285px;
            height: 34px;
            border-radius: 999px;
            accent-color: var(--blue);
            outline: none;
            display: block;
        }

        /* ---------- Messages ---------- */
        .empty {
            margin-top: 10px;
            padding: 54px 26px;
            text-align: center;
            border: 1px dashed var(--line-strong);
            border-radius: 20px;
            background: rgba(255, 255, 255, 0.02);
        }

        .empty .icon {
            font-size: 34px;
            display: block;
            margin-bottom: 10px;
            opacity: 0.75;
        }

        .empty h3 {
            margin: 0 0 6px;
            font-size: 17px;
        }

        .empty p {
            margin: 0;
            color: var(--muted);
            font-size: 13px;
        }

        .alert {
            display: block;
            margin-top: 20px;
            padding: 13px 18px;
            border-radius: 12px;
            font-size: 13px;
            color: #ffd7d7;
            background: rgba(214, 69, 69, 0.14);
            border: 1px solid rgba(214, 69, 69, 0.38);
        }

        .foot-note {
            margin-top: 26px;
            font-size: 12px;
            color: var(--muted);
            text-align: center;
        }

        @media (max-width: 900px) {
            .col-path, table.songs thead th:nth-child(3) {
                display: none;
            }

            .col-play {
                width: auto;
            }
        }

        @media (max-width: 640px) {
            .page {
                padding: 26px 14px 50px;
            }

            .page-head h1 {
                font-size: 26px;
            }

            .player {
                max-width: 100%;
            }
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <header class="app-bar">
            <div class="app-bar-inner">
                <div class="brand">
                    <span class="brand-mark">&#9835;</span>
                    <span>PoliMusic<br /><span class="brand-sub">Music library</span></span>
                </div>
                <div class="user-area">
                    <asp:Image ID="imgUserPhoto" runat="server" CssClass="avatar" AlternateText="User photo" />
                    <span class="welcome">Welcome
                        <asp:Label ID="lblUserName" runat="server" CssClass="user-name"></asp:Label></span>
                    <a class="logout" href="Logout.aspx">Logout</a>
                </div>
            </div>
        </header>

        <main class="page">
            <div class="page-head">
                <div>
                    <h1>Song <span class="accent">list</span></h1>
                    <p class="subtitle">Listen to the songs available in your PoliMusic library.</p>
                </div>
                <div class="counter">
                    <b><asp:Label ID="lblCount" runat="server" Text="0"></asp:Label></b>
                    <span>songs</span>
                </div>
            </div>

            <asp:Panel ID="pnlSongs" runat="server" CssClass="table-card">
                <table class="songs">
                    <thead>
                        <tr>
                            <th class="col-id">ID Canción</th>
                            <th>Nombre</th>
                            <th class="col-path">Ruta</th>
                            <th class="col-plays">Reproducciones</th>
                            <th class="col-play">Reproducir</th>
                        </tr>
                    </thead>
                    <tbody>
                        <asp:Repeater ID="rptSongs" runat="server">
                            <ItemTemplate>
                                <tr>
                                    <td class="col-id">
                                        <span class="id-pill"><%# Eval("ID") %></span>
                                    </td>
                                    <td>
                                        <span class="song-name"><%# Text(Eval("SongName")) %></span>
                                    </td>
                                    <td class="col-path">
                                        <span class="path"><%# Text(Eval("SongPath")) %></span>
                                    </td>
                                    <td class="col-plays">
                                        <span class="plays" id='plays-<%# Eval("ID") %>'><%# Eval("Plays") %></span>
                                    </td>
                                    <td class="col-play">
                                        <audio class="player" controls="controls" preload="none"
                                            data-song='<%# Eval("ID") %>'
                                            src='<%# GetSongUrl(Eval("SongPath")) %>'></audio>
                                    </td>
                                </tr>
                            </ItemTemplate>
                        </asp:Repeater>
                    </tbody>
                </table>
            </asp:Panel>

            <asp:Panel ID="pnlEmpty" runat="server" CssClass="empty" Visible="false">
                <span class="icon">&#9834;</span>
                <h3>There are no songs yet</h3>
                <p>The song list will be displayed here as soon as songs are registered.</p>
            </asp:Panel>

            <asp:Label ID="lblMessage" runat="server" CssClass="alert" Visible="false"></asp:Label>

            <p class="foot-note">PoliMusic &middot; EPN FIS &middot; Web Application</p>
        </main>

        <script type="text/javascript">
            (function () {
                // Only one song can be playing at the same time.
                function pauseOthers(current) {
                    var players = document.getElementsByTagName('audio');
                    for (var i = 0; i < players.length; i++) {
                        if (players[i] !== current) {
                            players[i].pause();
                        }
                    }
                }

                // Registers the reproduction on the server and refreshes the counter.
                function registerPlay(player) {
                    if (player.getAttribute('data-counted') === '1') {
                        return;
                    }
                    player.setAttribute('data-counted', '1');

                    var idSong = player.getAttribute('data-song');
                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', 'SongPlay.ashx?id=' + encodeURIComponent(idSong), true);
                    xhr.onreadystatechange = function () {
                        if (xhr.readyState !== 4 || xhr.status !== 200) {
                            return;
                        }
                        var badge = document.getElementById('plays-' + idSong);
                        if (!badge) {
                            return;
                        }
                        try {
                            var data = JSON.parse(xhr.responseText);
                            if (data && typeof data.plays !== 'undefined') {
                                badge.innerHTML = data.plays;
                            }
                        } catch (e) {
                            // The counter keeps its previous value.
                        }
                    };
                    xhr.send();
                }

                var players = document.getElementsByTagName('audio');
                for (var i = 0; i < players.length; i++) {
                    players[i].addEventListener('play', function () {
                        pauseOthers(this);
                        registerPlay(this);
                    });
                }
            })();
        </script>
    </form>
</body>
</html>
