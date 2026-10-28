<%
' ============================================================================
'  PoliMusic - Pie comun (cierre de main + barra del reproductor)
' ============================================================================
Dim ftrNoPlayer
ftrNoPlayer = (NoPlayer = True)
%>
    </div><!-- /.main-inner -->
  </main>

<% If Not ftrNoPlayer Then %>
  <!-- ======================= Reproductor ======================= -->
  <div class="player">
    <div class="player-now">
      <div class="np-art" id="npArt"></div>
      <div class="np-text">
        <div class="np-title" id="npTitle">Nada en reproduccion</div>
        <div class="np-artist" id="npArtist">Elige una cancion de la lista</div>
      </div>
    </div>

    <div class="player-center">
      <div class="player-controls">
        <button type="button" class="pc-btn" id="btnShuffle" title="Aleatorio"><%=Ico("shuffle", 16)%></button>
        <button type="button" class="pc-btn" id="btnPrev" title="Anterior"><%=Ico("prev", 16)%></button>
        <button type="button" class="pc-btn main" id="btnPlay" title="Reproducir / Pausar"><%=Ico("play", 14)%></button>
        <button type="button" class="pc-btn" id="btnNext" title="Siguiente"><%=Ico("next", 16)%></button>
        <button type="button" class="pc-btn" id="btnRepeat" title="Repetir"><%=Ico("repeat", 16)%></button>
      </div>
      <div class="player-progress">
        <span class="time" id="tCur">0:00</span>
        <div class="seek" id="seek"><div class="fill" id="seekFill"><span class="knob"></span></div></div>
        <span class="time" id="tDur">0:00</span>
      </div>
    </div>

    <div class="player-right">
      <button type="button" class="pc-btn" id="btnMute" title="Silenciar"><%=Ico("volume", 17)%></button>
      <input type="range" class="vol" id="vol" min="0" max="1" step="0.01" value="0.8" title="Volumen">
    </div>

    <audio id="audio" preload="none"></audio>
  </div>
<% End If %>

</div><!-- /.app -->
<script src="<%=H(AppRoot())%>assets/js/app.js"></script>
</body>
</html>
<%
' Cierre de la conexion con la base de datos
If Not IsEmpty(gConn) Then
    On Error Resume Next
    gConn.Close
    Set gConn = Nothing
    On Error GoTo 0
End If
%>
