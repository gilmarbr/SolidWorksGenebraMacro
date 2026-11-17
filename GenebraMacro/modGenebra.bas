Attribute VB_Name = "modGenebra"
'==============================================================================
' GERADOR DE MECANISMO GENEBRA EXTERNO  -  SOLIDWORKS 2024 (VBA)
' Autor: Gilmar Barbosa  -  Autom
'
' Cria: Roda Genebra, Manivela (placa + disco de travamento + pino(s)),
'       Base (ou peca so de referencia), eixos opcionais e a Montagem
'       com mates (base fixa; roda e manivela giram livres em seus eixos).
' Sincronismo: animacao cinematica exata pela macro (aba Movimento) e
'       roteiro para Estudo de Movimento (Motion) com motor + contato.
'
' Convencoes (montagem):
'   - Centro da roda na origem; centro da manivela em (C, 0).
'   - Angulo da manivela theta medido a partir de +X, anti-horario (+Z).
'   - theta = 0: pino apontando para FORA da roda = meio da parada.
'   - Unidades da interface: mm e graus. A API usa metros/radianos.
'==============================================================================
Option Explicit

Public swApp As SldWorks.SldWorks
Public Prm As Object                 ' Scripting.Dictionary com os parametros
Public G_PararAnim As Boolean
Public G_Acao As String              ' acao pedida pelo form: gerar / posicionar / animar
Public G_Resultado As String         ' texto mostrado na aba Resultados ao reabrir o form
Public G_ThetaPos As Double          ' angulo da manivela escolhido na aba Movimento

#If VBA7 Then
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If
Public G_Log As String

Public Const PI As Double = 3.14159265358979
Private Const MM As Double = 0.001
Private Const NOME_EIXO As String = "EIXO_GEN"

' ---- resultados do calculo (mm / graus) ----
Public G_n As Long, G_pinos As Long
Public G_beta As Double, G_bd As Double
Public G_C As Double, G_rk As Double, G_Re As Double, G_w As Double, G_dp As Double
Public G_Ro As Double, G_RoMin As Double, G_rhoB As Double
Public G_Rs As Double, G_Rlk As Double, G_Rr As Double, G_xi As Double, G_eP As Double
Public G_tR As Double, G_tM As Double, G_tL As Double, G_hp As Double, G_g As Double
Public G_Rplate As Double, G_bArm As Double, G_zB As Double
Public G_xmin As Double, G_xmax As Double, G_ymin As Double, G_ymax As Double
Public G_E As Double, G_frMov As Double, G_tIdx As Double, G_tPar As Double
Public G_w1 As Double, G_w2max As Double, G_a2max As Double, G_ratioMax As Double

Private G_PlFront As String
Private G_RotSign As Double

' ---- estado da animacao ----
Private A_md As Object, A_cR As Object, A_cM As Object

'==============================================================================
' PONTO DE ENTRADA
'==============================================================================
' O form e MODAL e so coleta parametros. As chamadas da API rodam aqui,
' fora dos eventos do form (chamar a API de dentro de eventos de form
' pode corromper a pilha do SOLIDWORKS e derrubar o programa).
Sub main()
    Set swApp = Application.SldWorks
    frmGenebra.Show vbModeless
End Sub

'==============================================================================
' ACESSO AOS PARAMETROS
'==============================================================================
Public Function PN(k As String) As Double
    If Prm Is Nothing Then Exit Function
    If Prm.Exists(k) Then PN = CDbl(Prm(k))
End Function

Public Function PS(k As String) As String
    If Prm Is Nothing Then Exit Function
    If Prm.Exists(k & "$") Then PS = CStr(Prm(k & "$"))
End Function

Public Function PB(k As String) As Boolean
    PB = (PN(k) <> 0)
End Function

' Decodifica \uXXXX em caracteres Unicode (os acentos ficam fora do arquivo .bas,
' que assim e 100% ASCII e nao depende da codificacao ao importar/colar).
Public Function U(ByVal s As String) As String
    Dim i As Long
    i = InStr(1, s, "\u")
    Do While i > 0
        s = Left$(s, i - 1) & ChrW$(CLng("&H" & Mid$(s, i + 2, 4))) & Mid$(s, i + 6)
        i = InStr(i + 1, s, "\u")
    Loop
    U = s
End Function

Private Function Mx(a As Double, b As Double) As Double
    If a > b Then Mx = a Else Mx = b
End Function

Private Function Mn(a As Double, b As Double) As Double
    If a < b Then Mn = a Else Mn = b
End Function

Private Function F2(x As Double) As String
    F2 = Format(x, "0.00")
End Function

Private Sub Erro(msgs As Collection, ByRef ok As Boolean, txt As String)
    msgs.Add "ERRO:  " & txt
    ok = False
End Sub

Private Sub Aviso(msgs As Collection, txt As String)
    msgs.Add "AVISO: " & txt
End Sub

Private Sub Passo(s As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ("TEMP") & "\Genebra_log.txt" For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & "  " & s
    Close #f
End Sub

Private Sub AddLog(s As String)
    G_Log = G_Log & s & vbCrLf
End Sub

Public Function JuntarMsgs(msgs As Collection) As String
    Dim i As Long, s As String
    For i = 1 To msgs.Count
        s = s & msgs(i) & vbCrLf
    Next
    JuntarMsgs = s
End Function

Public Function Atan2(y As Double, x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atan2 = Atn(y / x) + PI Else Atan2 = Atn(y / x) - PI
    Else
        If y > 0 Then
            Atan2 = PI / 2
        ElseIf y < 0 Then
            Atan2 = -PI / 2
        Else
            Atan2 = 0
        End If
    End If
End Function

Private Sub RotP(x As Double, y As Double, a As Double, ByRef xo As Double, ByRef yo As Double)
    xo = x * Cos(a) - y * Sin(a)
    yo = x * Sin(a) + y * Cos(a)
End Sub

'==============================================================================
' CINEMATICA
'==============================================================================
' Angulo (graus) do pino visto do centro da roda, para manivela em theta (graus)
Private Function AnguloPino(thDeg As Double) As Double
    Dim t As Double
    t = thDeg * PI / 180#
    AnguloPino = Atan2(G_rk * Sin(t), G_C + G_rk * Cos(t)) * 180# / PI
End Function

' Angulo acumulado da roda (graus, +Z) para um angulo da manivela theta (graus)
Public Function AnguloRoda(thDeg As Double) As Double
    Dim per As Double, u As Double, k As Double, f As Double, E As Double, d As Double
    per = 360# / G_pinos
    u = thDeg - (90# + G_bd)
    k = Int(u / per)
    f = u - k * per
    E = 180# - 2# * G_bd
    If f <= E Then
        d = AnguloPino(90# + G_bd + f) - G_bd
    Else
        d = -2# * G_bd
    End If
    AnguloRoda = -k * 2# * G_bd + d
End Function

Public Sub CarregarCinematica(n As Long, C As Double, pinos As Long)
    G_n = n: G_C = C
    If pinos >= 2 Then G_pinos = 2 Else G_pinos = 1
    G_beta = PI / G_n: G_bd = 180# / G_n
    G_rk = G_C * Sin(G_beta)
End Sub

'==============================================================================
' CALCULO E VALIDACAO
'==============================================================================
Public Function Calcular(msgs As Collection) As Boolean
    Dim ok As Boolean, ref As Double, fs As Double, ffundo As Double
    Dim fl As Double, ep As Double, fa As Double
    Dim dFR As Double, dFM As Double, dCR As Double, dCM As Double
    Dim hCR As Double, hCM As Double, kR As Double
    Dim fundo As Double, wallM As Double, Rk As Double, mB As Double
    Dim rpm As Double, modo As Long, aK As Double, i As Integer
    Dim xc As Double, yc As Double, x1 As Double, y1 As Double, cxr As Double
    Dim th As Double, h As Double, p0 As Double, p1 As Double, p2 As Double
    Dim d1 As Double, d2 As Double, m1 As Double, m2 As Double

    On Error GoTo Falha
    ok = True

    ' ---------- geral ----------
    G_n = CLng(PN("n"))
    If PN("pinos") >= 2 Then G_pinos = 2 Else G_pinos = 1
    If G_n < 3 Or G_n > 18 Then
        Call Erro(msgs, ok, U("N\u00BA de ranhuras deve estar entre 3 e 18."))
        Calcular = False: Exit Function
    End If
    G_beta = PI / G_n: G_bd = 180# / G_n
    G_dp = PN("dp"): fs = PN("fs")
    If G_dp <= 0 Then
        Call Erro(msgs, ok, U("\u00D8 do pino deve ser maior que zero."))
        Calcular = False: Exit Function
    End If
    If fs < 0 Then Call Erro(msgs, ok, U("A folga pino/ranhura n\u00E3o pode ser negativa."))
    G_w = G_dp + fs
    G_g = PN("gap")
    If G_g < 0 Then Call Erro(msgs, ok, U("A folga axial n\u00E3o pode ser negativa."))

    ref = PN("ref"): modo = CLng(PN("modo"))
    Select Case modo
        Case 1
            G_C = ref / Sin(G_beta)
        Case 2
            If ref / 2 <= G_w / 2 Then
                Call Erro(msgs, ok, U("Di\u00E2metro externo da roda muito pequeno."))
                Calcular = False: Exit Function
            End If
            G_C = Sqr((ref / 2) ^ 2 - (G_w / 2) ^ 2) / Cos(G_beta)
        Case Else
            G_C = ref
    End Select
    If G_C <= 0 Then
        Call Erro(msgs, ok, U("O valor de refer\u00EAncia (C, r ou \u00D8) deve ser maior que zero."))
        Calcular = False: Exit Function
    End If

    ' ---------- roda ----------
    G_rk = G_C * Sin(G_beta)
    G_Re = G_C * Cos(G_beta)
    G_RoMin = Sqr(G_Re ^ 2 + (G_w / 2) ^ 2)
    If modo = 2 Then
        G_Ro = ref / 2
    ElseIf PN("RoMan") > 0 Then
        G_Ro = PN("RoMan")
    Else
        G_Ro = G_RoMin
    End If
    If G_Ro < G_RoMin - 0.001 Then Call Erro(msgs, ok, "Raio externo da roda (" & F2(G_Ro) & U(") menor que o m\u00EDnimo (") & F2(G_RoMin) & ").")

    G_tR = PN("tR")
    If G_tR <= 0 Then Call Erro(msgs, ok, "Espessura da roda deve ser maior que zero.")
    ffundo = PN("ffundo")
    ' centro do semicirculo do fundo da ranhura
    G_rhoB = G_C - G_rk - G_dp / 2 + G_w / 2 - ffundo
    fundo = G_rhoB - G_w / 2
    If fundo <= 0 Then Call Erro(msgs, ok, "Fundo da ranhura passa do centro da roda.")

    ' Ponta do garfo: parede entre a lateral da ranhura e o arco de travamento.
    ' O ponto mais estreito fica no raio ~Re (na propria ponta) e vale
    ' (rk - w/2) - Rlk. Com Rs automatico, Rlk e escolhido para dar ePonta.
    fl = PN("fl"): ep = PN("ePonta"): fa = PN("fa")
    If fl < 0 Then Call Erro(msgs, ok, U("A folga do arco de travamento n\u00E3o pode ser negativa."))
    If PN("RsMan") > 0 Then
        G_Rs = PN("RsMan")
    Else
        If ep <= 0 Then Call Erro(msgs, ok, "Espessura da ponta do garfo deve ser maior que zero.")
        G_Rs = G_rk - G_w / 2 - ep - fl
    End If
    If G_Rs <= 0 Then
        Call Erro(msgs, ok, U("Raio do disco de travamento (Rs) resultou <= 0. Reduza a espessura da ponta, a folga do arco ou o pino."))
        Calcular = False: Exit Function
    End If
    G_Rlk = G_Rs + fl
    G_eP = (G_rk - G_w / 2) - G_Rlk
    If G_eP <= 0.0001 Then
        Call Erro(msgs, ok, "Arco de travamento (R=" & F2(G_Rlk) & U(") invade as ranhuras. Reduza Rs ou a folga do arco."))
    ElseIf PN("RsMan") > 0 And ep > 0 And G_eP < ep - 0.001 Then
        Call Aviso(msgs, U("Com o Rs manual, a ponta do garfo fica com ") & F2(G_eP) & U(" mm (pedido: ") & F2(ep) & " mm).")
    ElseIf G_eP < 0.5 Then
        Call Aviso(msgs, U("Ponta do garfo com menos de 0,5 mm (") & F2(G_eP) & " mm).")
    End If

    dFR = PN("dFuroR")
    If dFR <= 0 Then Call Erro(msgs, ok, U("\u00D8 do furo da roda deve ser maior que zero."))
    If PB("chavR") Then
        kR = PN("tChR")
        If PN("bChR") >= dFR Then Call Erro(msgs, ok, U("Largura da chaveta da roda >= \u00D8 do furo."))
    End If
    If dFR / 2 + kR > fundo - 2 Then Call Erro(msgs, ok, "Furo/chaveta da roda muito perto do fundo das ranhuras (fundo em R=" & F2(fundo) & ").")
    If PB("cuboR") Then
        dCR = PN("dCuboR"): hCR = PN("hCuboR")
        If dCR <= dFR + 2 * kR + 2 Then Call Erro(msgs, ok, U("\u00D8 do cubo da roda deve ser maior que o furo + chaveta."))
        If dCR / 2 > Mn(fundo, G_C - G_Rlk) - 1 Then Call Erro(msgs, ok, "Cubo da roda invade ranhuras ou arcos de travamento.")
        If hCR <= 0 Then Call Erro(msgs, ok, "Altura do cubo da roda deve ser maior que zero.")
    End If
    If G_rhoB * Sin(G_beta) - G_w / 2 < 1 Then Call Erro(msgs, ok, "Ranhuras vizinhas se encontram perto do centro. Reduza o pino ou aumente C.")

    ' ---------- manivela ----------
    G_tM = PN("tM")
    If G_tM <= 0 Then Call Erro(msgs, ok, "Espessura da placa da manivela deve ser maior que zero.")
    If PN("tL") > 0 Then G_tL = PN("tL") Else G_tL = G_tR + G_g
    If PN("hp") > 0 Then G_hp = PN("hp") Else G_hp = G_g + G_tR + 1
    If G_hp < G_g + G_tR Then Call Erro(msgs, ok, "Pino curto: precisa de pelo menos " & F2(G_g + G_tR) & " mm para atravessar a roda.")

    G_Rr = G_Ro + fa
    G_xi = (G_Rs ^ 2 - G_Rr ^ 2 + G_C ^ 2) / (2 * G_C)
    If Abs(G_xi) >= G_Rs Then
        Call Erro(msgs, ok, U("O al\u00EDvio n\u00E3o intercepta o disco de travamento. Verifique Rs e a folga do al\u00EDvio."))
        Calcular = False: Exit Function
    End If
    If G_pinos = 2 And G_xi <= 0.5 Then Call Erro(msgs, ok, U("Com 2 pinos os al\u00EDvios se sobrep\u00F5em no centro. Aumente C ou reduza Rs."))

    dFM = PN("dFuroM")
    If dFM <= 0 Then Call Erro(msgs, ok, U("\u00D8 do furo da manivela deve ser maior que zero."))
    wallM = G_C - G_Rr - dFM / 2
    If wallM < 1 Then Call Erro(msgs, ok, U("O al\u00EDvio do disco de travamento invade o furo da manivela (parede ") & F2(wallM) & U(" mm). Reduza o \u00D8 do eixo, aumente C ou use menos ranhuras."))
    If PB("chavM") Then
        If PN("bChM") >= dFM Then Call Erro(msgs, ok, U("Largura da chaveta da manivela >= \u00D8 do furo."))
        If G_pinos = 1 Then aK = 180# Else aK = 90#
        For i = -1 To 1 Step 2
            Call RotP(i * PN("bChM") / 2, dFM / 2 + PN("tChM"), (aK - 90#) * PI / 180#, x1, y1)
            If Sqr(x1 ^ 2 + y1 ^ 2) > G_Rs - 1 Then Call Erro(msgs, ok, "Chaveta da manivela sai do disco de travamento.")
            cxr = G_C
            If Sqr((x1 - cxr) ^ 2 + y1 ^ 2) < G_Rr + 0.5 Then Call Erro(msgs, ok, U("Chaveta da manivela invade o al\u00EDvio."))
            If G_pinos = 2 Then
                If Sqr((x1 + cxr) ^ 2 + y1 ^ 2) < G_Rr + 0.5 Then Call Erro(msgs, ok, U("Chaveta da manivela invade o al\u00EDvio."))
            End If
        Next
    End If
    If dFM / 2 > G_Rs - 1 Then Call Erro(msgs, ok, "Furo da manivela maior que o disco de travamento.")

    If PB("cuboM") Then
        dCM = PN("dCuboM"): hCM = PN("hCuboM")
        If dCM <= dFM + 2 Then Call Erro(msgs, ok, U("\u00D8 do cubo da manivela deve ser maior que o furo."))
        If hCM <= 0 Then Call Erro(msgs, ok, "Altura do cubo da manivela deve ser maior que zero.")
    End If

    If CLng(PN("forma")) = 1 Then
        If PN("bBraco") > 0 Then
            G_bArm = PN("bBraco")
        Else
            G_bArm = Mx(G_dp, Mx(dFM, dCM)) + 2 * Mx(3, G_dp / 2)
        End If
        If G_bArm < G_dp + 2 Then Call Erro(msgs, ok, U("Bra\u00E7o estreito demais para o pino."))
        If G_bArm < dFM + 4 Then Call Erro(msgs, ok, U("Bra\u00E7o estreito demais para o furo."))
        G_Rplate = G_rk + G_bArm / 2
    Else
        If PN("RmMan") > 0 Then G_Rplate = PN("RmMan") Else G_Rplate = G_rk + G_dp / 2 + Mx(3, G_dp / 2)
        If G_Rplate < G_rk + G_dp / 2 + 1 Then Call Erro(msgs, ok, "Disco da manivela pequeno demais para o pino.")
    End If

    ' ---------- interferencias entre conjuntos ----------
    If G_C - G_Rplate - dFR / 2 < 1 Then Call Erro(msgs, ok, U("Placa/bra\u00E7o da manivela bate no eixo da roda. Use 'Bra\u00E7o', reduza o pino ou aumente C."))
    If dCM > 0 Then
        If G_C - dCM / 2 - dFR / 2 < 1 Then Call Erro(msgs, ok, "Cubo da manivela bate no eixo da roda.")
    End If
    If dCR > 0 Then
        If G_C - dCR / 2 - dFM / 2 < 1 Then Call Erro(msgs, ok, "Cubo da roda bate no eixo da manivela.")
        If G_C - G_rk - G_dp / 2 - dCR / 2 < 1 Then Call Erro(msgs, ok, "Pino bate no cubo da roda.")
    End If

    ' ---------- base ----------
    G_zB = -(G_g + G_tM + hCM + G_g)
    mB = PN("mB")
    Rk = Mx(G_Rplate, Mx(G_Rs, dCM / 2))
    G_xmin = -G_Ro - mB
    G_xmax = G_C + Rk + mB
    G_ymax = Mx(G_Ro, Rk) + mB
    G_ymin = -G_ymax
    If PB("gBase") Then
        If PN("tB") <= 0 Then Call Erro(msgs, ok, "Espessura da base deve ser maior que zero.")
        If PN("nFix") >= 4 And PN("dFix") > mB - 2 Then Call Erro(msgs, ok, U("Furos de fixa\u00E7\u00E3o maiores que a margem da base."))
        If PN("rcB") > mB Then Call Aviso(msgs, U("Raio dos cantos maior que a margem: os furos de fixa\u00E7\u00E3o podem ficar fora da base."))
    End If

    ' ---------- cinematica ----------
    rpm = PN("rpm"): If rpm <= 0 Then rpm = 60
    G_E = 180# - 2# * G_bd
    G_frMov = G_E * G_pinos / 360#
    G_w1 = rpm * 2# * PI / 60#
    G_tIdx = G_E / 360# * 60# / rpm
    G_tPar = (360# / G_pinos - G_E) / 360# * 60# / rpm
    h = 0.05: m1 = 0: m2 = 0
    For th = 90# + G_bd + h To 270# - G_bd - h Step h
        p0 = AnguloPino(th - h): p1 = AnguloPino(th): p2 = AnguloPino(th + h)
        d1 = Abs((p2 - p0) / (2 * h))
        d2 = Abs((p2 - 2 * p1 + p0) / (h * h)) * 180# / PI
        If d1 > m1 Then m1 = d1
        If d2 > m2 Then m2 = d2
    Next
    G_ratioMax = m1
    G_w2max = m1 * G_w1
    G_a2max = m2 * G_w1 ^ 2
    If G_n = 3 Then Call Aviso(msgs, U("n = 3 gera acelera\u00E7\u00F5es muito altas na roda."))

    Calcular = ok
    Exit Function
Falha:
    msgs.Add U("ERRO:  valores inv\u00E1lidos (") & Err.Description & ")."
    Calcular = False
End Function

Public Function TextoResultados(ok As Boolean, msgs As Collection) As String
    Dim s As String
    s = "=== GEOMETRIA (mm) ===" & vbCrLf
    s = s & U("Dist\u00E2ncia entre centros C ........ ") & F2(G_C) & vbCrLf
    s = s & "Raio do pino na manivela r ....... " & F2(G_rk) & vbCrLf
    s = s & "Raio de entrada na roda R ........ " & F2(G_Re) & vbCrLf
    s = s & "Raio externo da roda Ro .......... " & F2(G_Ro) & U("   (m\u00EDnimo ") & F2(G_RoMin) & ")" & vbCrLf
    s = s & U("\u00D8 externo da roda ................ ") & F2(2 * G_Ro) & vbCrLf
    s = s & "Largura da ranhura ............... " & F2(G_w) & vbCrLf
    s = s & "Fundo da ranhura (raio) .......... " & F2(G_rhoB - G_w / 2) & vbCrLf
    s = s & "Raio do disco de travamento Rs ... " & F2(G_Rs) & vbCrLf
    s = s & "Raio do arco de travamento (roda)  " & F2(G_Rlk) & vbCrLf
    s = s & "Espessura da ponta do garfo ...... " & F2(G_eP) & vbCrLf
    s = s & U("Raio do al\u00EDvio (crescente) ....... ") & F2(G_Rr) & vbCrLf
    s = s & "Esp. do disco de travamento ...... " & F2(G_tL) & vbCrLf
    s = s & "Altura do pino ................... " & F2(G_hp) & vbCrLf
    s = s & U("Raio externo da placa/bra\u00E7o ...... ") & F2(G_Rplate) & vbCrLf
    s = s & "Base (C x L) ..................... " & F2(G_xmax - G_xmin) & " x " & F2(G_ymax - G_ymin) & vbCrLf
    s = s & vbCrLf & U("=== CINEM\u00C1TICA ===") & vbCrLf
    s = s & U("Indexa\u00E7\u00E3o da roda por ciclo ...... ") & F2(2 * G_bd) & U("\u00B0") & vbCrLf
    s = s & U("\u00C2ngulo de engate da manivela ..... ") & F2(G_E) & U("\u00B0 (por pino)") & vbCrLf
    s = s & "Movimento / parada ............... " & Format(G_frMov * 100, "0.0") & "% / " & Format((1 - G_frMov) * 100, "0.0") & "%" & vbCrLf
    s = s & U("Tempo de indexa\u00E7\u00E3o ............... ") & Format(G_tIdx, "0.000") & " s" & vbCrLf
    s = s & "Tempo de parada .................. " & Format(G_tPar, "0.000") & " s" & vbCrLf
    s = s & U("w2/w1 m\u00E1ximo ..................... ") & Format(G_ratioMax, "0.000") & vbCrLf
    s = s & U("w2 m\u00E1ximo ........................ ") & Format(G_w2max, "0.00") & " rad/s  (" & Format(G_w2max * 60 / (2 * PI), "0.0") & " rpm)" & vbCrLf
    s = s & U("Acelera\u00E7\u00E3o angular m\u00E1x. da roda .. ") & Format(G_a2max, "0.0") & U(" rad/s\u00B2") & vbCrLf
    s = s & vbCrLf & U("=== VALIDA\u00C7\u00C3O ===") & vbCrLf
    If msgs.Count = 0 Then
        s = s & "Nenhum problema encontrado." & vbCrLf
    Else
        s = s & JuntarMsgs(msgs)
    End If
    If ok Then s = s & vbCrLf & ">> Pronto para gerar." Else s = s & vbCrLf & ">> Corrija os erros antes de gerar."
    TextoResultados = s
End Function

'==============================================================================
' GERACAO
'==============================================================================
Public Function Gerar() As String
    Dim msgs As New Collection, pasta As String, pre As String
    Dim aR As String, aM As String, aB As String, aER As String, aEM As String, aA As String
    Dim topoR As Double, topoM As Double, tB As Double

    G_Log = ""
    If swApp Is Nothing Then Set swApp = Application.SldWorks
    If Not Calcular(msgs) Then
        Gerar = "Corrija os erros antes de gerar:" & vbCrLf & JuntarMsgs(msgs)
        Exit Function
    End If

    On Error GoTo Falha
    pasta = Trim(PS("pasta"))
    If pasta = "" Then Err.Raise vbObjectError + 1, , U("Informe a pasta de sa\u00EDda.")
    If Right(pasta, 1) = "\" Then pasta = Left(pasta, Len(pasta) - 1)
    CriarPasta pasta

    pre = Trim(PS("prefixo"))
    If pre = "" Then pre = "GEN_n" & G_n & "_C" & Format(G_C, "0") & "_"
    aR = pasta & "\" & pre & "Roda.SLDPRT"
    aM = pasta & "\" & pre & "Manivela.SLDPRT"
    aB = pasta & "\" & pre & "Base.SLDPRT"
    aER = pasta & "\" & pre & "EixoRoda.SLDPRT"
    aEM = pasta & "\" & pre & "EixoManivela.SLDPRT"
    aA = pasta & "\" & pre & "Montagem.SLDASM"

    If Not PB("sobrescrever") Then
        If Dir(aR) <> "" Or Dir(aM) <> "" Or Dir(aB) <> "" Or Dir(aA) <> "" Then
            Err.Raise vbObjectError + 2, , U("J\u00E1 existem arquivos com o prefixo '") & pre & "'. Marque 'Sobrescrever' ou mude o prefixo."
        End If
    End If

    FecharDoc aA: FecharDoc aR: FecharDoc aM: FecharDoc aB: FecharDoc aER: FecharDoc aEM

    On Error Resume Next: Kill Environ("TEMP") & "\Genebra_log.txt": On Error GoTo Falha
    Passo "=== INICIO GERAR ==="
    G_RotSign = 0
    DetectarSinalRotacao

    CriarRoda aR
    CriarManivela aM
    CriarBase aB
    If PB("gEixos") Then
        If PB("gBase") Then tB = PN("tB") Else tB = 0
        topoR = G_tR + PN("sobra") - G_zB
        If PB("cuboR") Then topoR = topoR + PN("hCuboR")
        topoM = -G_g + G_tL + PN("sobra") - G_zB
        CriarEixo aER, PN("dFuroR"), topoR, tB, PS("matE"), CorEixo()
        CriarEixo aEM, PN("dFuroM"), topoM, tB, PS("matE"), CorEixo()
    End If
    CriarMontagem aA, aR, aM, aB, aER, aEM

    If PB("gCSV") Then GravarCSV pasta & "\" & pre & "Cinematica.csv"
    GravarMotion pasta & "\" & pre & "LEIA-ME_Motion.txt", aA

    Gerar = U("CONCLU\u00CDDO") & vbCrLf & vbCrLf & G_Log & vbCrLf & _
            "Para o Estudo de Movimento veja: " & pre & "LEIA-ME_Motion.txt" & vbCrLf & _
            "Para animar: aba 'Movimento' (com a montagem ativa)."
    Exit Function
Falha:
    Gerar = U("ERRO na gera\u00E7\u00E3o: ") & Err.Description & vbCrLf & vbCrLf & G_Log
End Function

'------------------------------------------------------------------------------
' RODA GENEBRA
'------------------------------------------------------------------------------
Private Sub CriarRoda(arq As String)
    Passo "### CriarRoda"
    Dim md As Object, sm As Object, pl As Variant, sk As Object, i As Long
    Dim psi As Double, ang As Double
    Set md = NovaPeca()
    Set sm = md.SketchManager
    pl = PlanosPadrao(md)

    ' 1) disco
    IniciarEsboco md, pl(0)
    Call Circulo(sm, 0, 0, G_Ro)
    Call DesenharFuro(sm, 0, 0, PN("dFuroR"), PB("chavR"), PN("bChR"), PN("tChR"), 90)
    Set sk = FecharEsboco(md)
    Extrudar md, sk, G_tR, False, "Disco da roda"

    ' 2) cubo (lado da frente, +Z)
    If PB("cuboR") Then
        IniciarEsboco md, pl(0)
        Call Circulo(sm, 0, 0, PN("dCuboR") / 2)
        Call DesenharFuro(sm, 0, 0, PN("dFuroR"), PB("chavR"), PN("bChR"), PN("tChR"), 90)
        Set sk = FecharEsboco(md)
        Extrudar md, sk, G_tR + PN("hCuboR"), False, "Cubo da roda"
    End If

    ' 3) ranhuras + arcos de travamento (contornos separados, corte passante)
    IniciarEsboco md, pl(0)
    For i = 0 To G_n - 1
        psi = G_beta + 2# * G_beta * i
        Call DesenharRanhura(sm, psi)
        ang = 2# * G_beta * i
        Call Circulo(sm, G_C * Cos(ang), G_C * Sin(ang), G_Rlk)
    Next
    Set sk = FecharEsboco(md)
    CortarPassante md, sk, "Ranhuras e travamento"

    EixoReferencia md, pl(1), pl(2), NOME_EIXO
    AplicarMaterialCor md, PS("matR"), PS("corR")
    md.ForceRebuild3 False
    Salvar md, arq
End Sub

Private Sub DesenharRanhura(sm As Object, psi As Double)
    Dim ux As Double, uy As Double, vx As Double, vy As Double, hw As Double, ro As Double
    ux = Cos(psi): uy = Sin(psi): vx = -uy: vy = ux
    hw = G_w / 2: ro = G_Ro + 2
    Call Linha(sm, G_rhoB * ux + hw * vx, G_rhoB * uy + hw * vy, ro * ux + hw * vx, ro * uy + hw * vy)
    Call Linha(sm, ro * ux + hw * vx, ro * uy + hw * vy, ro * ux - hw * vx, ro * uy - hw * vy)
    Call Linha(sm, ro * ux - hw * vx, ro * uy - hw * vy, G_rhoB * ux - hw * vx, G_rhoB * uy - hw * vy)
    Call Arco(sm, G_rhoB * ux, G_rhoB * uy, G_rhoB * ux - hw * vx, G_rhoB * uy - hw * vy, G_rhoB * ux + hw * vx, G_rhoB * uy + hw * vy, -1)
End Sub

'------------------------------------------------------------------------------
' MANIVELA (placa + cubo + disco de travamento + pinos)
'------------------------------------------------------------------------------
Private Sub CriarManivela(arq As String)
    Passo "### CriarManivela"
    Dim md As Object, sm As Object, pl As Variant, sk As Object, aK As Double
    Set md = NovaPeca()
    Set sm = md.SketchManager
    pl = PlanosPadrao(md)
    If G_pinos = 1 Then aK = 180# Else aK = 90#

    ' 1) placa / braco (atras, -Z)
    IniciarEsboco md, pl(0)
    If CLng(PN("forma")) = 1 Then
        DesenharBraco sm
    Else
        Call Circulo(sm, 0, 0, G_Rplate)
    End If
    Call DesenharFuro(sm, 0, 0, PN("dFuroM"), PB("chavM"), PN("bChM"), PN("tChM"), aK)
    Set sk = FecharEsboco(md)
    Extrudar md, sk, G_tM, True, "Placa da manivela"

    ' 2) cubo (atras da placa)
    If PB("cuboM") Then
        IniciarEsboco md, pl(0)
        Call Circulo(sm, 0, 0, PN("dCuboM") / 2)
        Call DesenharFuro(sm, 0, 0, PN("dFuroM"), PB("chavM"), PN("bChM"), PN("tChM"), aK)
        Set sk = FecharEsboco(md)
        Extrudar md, sk, G_tM + PN("hCuboM"), True, "Cubo da manivela"
    End If

    ' 3) disco de travamento com alivio (frente, +Z)
    IniciarEsboco md, pl(0)
    DesenharDiscoTravamento sm
    Call DesenharFuro(sm, 0, 0, PN("dFuroM"), PB("chavM"), PN("bChM"), PN("tChM"), aK)
    Set sk = FecharEsboco(md)
    Extrudar md, sk, G_tL, False, "Disco de travamento"

    ' 4) pino(s)
    IniciarEsboco md, pl(0)
    Call Circulo(sm, G_rk, 0, G_dp / 2)
    If G_pinos = 2 Then Call Circulo(sm, -G_rk, 0, G_dp / 2)
    Set sk = FecharEsboco(md)
    Extrudar md, sk, G_hp, False, "Pino"

    EixoReferencia md, pl(1), pl(2), NOME_EIXO
    AplicarMaterialCor md, PS("matM"), PS("corM")
    md.ForceRebuild3 False
    Salvar md, arq
End Sub

Private Sub DesenharBraco(sm As Object)
    Dim h As Double, xa As Double
    h = G_bArm / 2
    If G_pinos = 2 Then xa = -G_rk Else xa = 0
    Call Linha(sm, xa, -h, G_rk, -h)
    Call Arco(sm, G_rk, 0, G_rk, -h, G_rk, h, 1)
    Call Linha(sm, G_rk, h, xa, h)
    Call Arco(sm, xa, 0, xa, h, xa, -h, 1)
End Sub

Private Sub DesenharDiscoTravamento(sm As Object)
    Dim yi As Double
    yi = Sqr(G_Rs ^ 2 - G_xi ^ 2)
    If G_pinos = 1 Then
        Call Arco(sm, 0, 0, G_xi, yi, G_xi, -yi, 1)
        Call Arco(sm, G_C, 0, G_xi, -yi, G_xi, yi, -1)
    Else
        Call Arco(sm, 0, 0, G_xi, yi, -G_xi, yi, 1)
        Call Arco(sm, -G_C, 0, -G_xi, yi, -G_xi, -yi, -1)
        Call Arco(sm, 0, 0, -G_xi, -yi, G_xi, -yi, 1)
        Call Arco(sm, G_C, 0, G_xi, -yi, G_xi, yi, -1)
    End If
End Sub

'------------------------------------------------------------------------------
' BASE (ou peca so com a referencia) - contem eixos e planos usados nos mates
'------------------------------------------------------------------------------
Private Sub CriarBase(arq As String)
    Passo "### CriarBase"
    Dim md As Object, sm As Object, pl As Variant, sk As Object, pX As Object
    Dim m2 As Double
    Set md = NovaPeca()
    Set sm = md.SketchManager
    pl = PlanosPadrao(md)

    If PB("gBase") Then
        IniciarEsboco md, pl(0)
        Call RetArred(sm, G_xmin, G_ymin, G_xmax, G_ymax, PN("rcB"))
        Call Circulo(sm, 0, 0, PN("dFuroR") / 2)
        Call Circulo(sm, G_C, 0, PN("dFuroM") / 2)
        If PN("nFix") >= 4 Then
            m2 = PN("mB") / 2
            Call Circulo(sm, G_xmin + m2, G_ymin + m2, PN("dFix") / 2)
            Call Circulo(sm, G_xmax - m2, G_ymin + m2, PN("dFix") / 2)
            Call Circulo(sm, G_xmax - m2, G_ymax - m2, PN("dFix") / 2)
            Call Circulo(sm, G_xmin + m2, G_ymax - m2, PN("dFix") / 2)
        End If
        Set sk = FecharEsboco(md)
        Extrudar md, sk, PN("tB"), True, "Placa base"
        AplicarMaterialCor md, PS("matB"), PS("corB")
    Else
        AddLog U("Base sem corpo (somente refer\u00EAncias).")
    End If

    EixoReferencia md, pl(1), pl(2), "EIXO_R"
    Set pX = PlanoOffset(md, pl(2), G_C, 0, "PL_X_MAN")
    EixoReferencia md, pl(1), pX, "EIXO_M"
    PlanoOffset md, pl(0), -G_zB, 2, "PL_RODA"
    PlanoOffset md, pl(0), -G_zB - G_g, 2, "PL_MAN"
    md.ForceRebuild3 False
    Salvar md, arq
End Sub

'------------------------------------------------------------------------------
' EIXO
'------------------------------------------------------------------------------
Private Sub CriarEixo(arq As String, d As Double, topo As Double, abaixo As Double, mat As String, cor As String)
    Passo "### CriarEixo"
    Dim md As Object, sm As Object, pl As Variant, sk As Object
    Set md = NovaPeca()
    Set sm = md.SketchManager
    pl = PlanosPadrao(md)
    IniciarEsboco md, pl(0)
    Call Circulo(sm, 0, 0, d / 2)
    Set sk = FecharEsboco(md)
    If abaixo > 0 Then
        Extrudar md, sk, topo, False, "Eixo", True, abaixo
    Else
        Extrudar md, sk, topo, False, "Eixo"
    End If
    EixoReferencia md, pl(1), pl(2), NOME_EIXO
    AplicarMaterialCor md, mat, cor
    md.ForceRebuild3 False
    Salvar md, arq
End Sub

'------------------------------------------------------------------------------
' MONTAGEM
'------------------------------------------------------------------------------
Private Sub CriarMontagem(aA As String, aR As String, aM As String, aB As String, aER As String, aEM As String)
    Passo "### CriarMontagem"
    Dim md As Object, asy As Object, cB As Object, cR As Object, cM As Object
    Dim cER As Object, cEM As Object, th0 As Double, cpm As Object

    Passo "Montagem: NewDocument"
    Set md = swApp.NewDocument(Template(9, "tplMont"), 0, 0, 0)
    Passo "Montagem: ok"
    If md Is Nothing Then Err.Raise vbObjectError + 30, , U("N\u00E3o foi poss\u00EDvel criar a montagem (verifique o template).")
    If md.GetType <> 2 Then Err.Raise vbObjectError + 31, , U("O template de montagem n\u00E3o \u00E9 um .asmdot.")
    Set asy = md

    Passo "AddComponent5"
    Set cB = asy.AddComponent5(aB, 0, "", False, "", 0, 0, G_zB * MM)
    If cB Is Nothing Then Err.Raise vbObjectError + 32, , "Falha ao inserir a base."
    md.ClearSelection2 True
    cB.Select4 False, Nothing, False
    Passo "FixComponent"
    asy.FixComponent
    md.ClearSelection2 True

    Passo "AddComponent5"
    Set cR = asy.AddComponent5(aR, 0, "", False, "", 0, 0, 0)
    Passo "AddComponent5"
    Set cM = asy.AddComponent5(aM, 0, "", False, "", G_C * MM, 0, -G_g * MM)
    If cR Is Nothing Or cM Is Nothing Then Err.Raise vbObjectError + 33, , "Falha ao inserir roda/manivela."

    th0 = PN("theta0")
    If Not AnguloZero(AnguloRoda(th0)) Then RotacionarZ cR, AnguloRoda(th0)
    If Not AnguloZero(th0) Then RotacionarZ cM, th0

    ' mates: base fixa; roda e manivela livres apenas na rotacao
    Acoplar md, asy, cR.FeatureByName(NOME_EIXO), cB.FeatureByName("EIXO_R"), 0, 2, "Eixo da roda"
    Acoplar md, asy, cR.FeatureByName(G_PlFront), cB.FeatureByName("PL_RODA"), 0, 2, "Plano da roda"
    Acoplar md, asy, cM.FeatureByName(NOME_EIXO), cB.FeatureByName("EIXO_M"), 0, 2, "Eixo da manivela"
    Acoplar md, asy, cM.FeatureByName(G_PlFront), cB.FeatureByName("PL_MAN"), 0, 2, "Plano da manivela"

    If PB("gEixos") Then
        Passo "AddComponent5"
        Set cER = asy.AddComponent5(aER, 0, "", False, "", 0, 0, G_zB * MM)
        Passo "AddComponent5"
        Set cEM = asy.AddComponent5(aEM, 0, "", False, "", G_C * MM, 0, G_zB * MM)
        If Not cER Is Nothing Then
            Acoplar md, asy, cER.FeatureByName(NOME_EIXO), cB.FeatureByName("EIXO_R"), 0, 2, U("Eixo (pe\u00E7a) roda")
            Acoplar md, asy, cER.FeatureByName(G_PlFront), cB.FeatureByName(G_PlFront), 0, 2, "Apoio eixo roda"
        End If
        If Not cEM Is Nothing Then
            Acoplar md, asy, cEM.FeatureByName(NOME_EIXO), cB.FeatureByName("EIXO_M"), 0, 2, U("Eixo (pe\u00E7a) manivela")
            Acoplar md, asy, cEM.FeatureByName(G_PlFront), cB.FeatureByName(G_PlFront), 0, 2, "Apoio eixo manivela"
        End If
    End If

    md.EditRebuild3
    If Not AnguloZero(AnguloRoda(th0)) Then RotacionarZ cR, AnguloRoda(th0)
    If Not AnguloZero(th0) Then RotacionarZ cM, th0
    md.EditRebuild3

    ' propriedades usadas pela animacao
    Set cpm = md.Extension.CustomPropertyManager("")
    Passo "Propriedades"
    cpm.Add3 "GEN_N", 30, CStr(G_n), 2
    cpm.Add3 "GEN_C", 30, Replace(CStr(G_C), ",", "."), 2
    cpm.Add3 "GEN_PINOS", 30, CStr(G_pinos), 2
    cpm.Add3 "GEN_RODA", 30, cR.Name2, 2
    cpm.Add3 "GEN_MANIVELA", 30, cM.Name2, 2

    Passo "Vista/zoom"
    md.ShowNamedView2 "", 7
    md.ViewZoomtofit2
    Salvar md, aA
End Sub

Private Sub Acoplar(md As Object, asy As Object, f1 As Object, f2 As Object, tipo As Long, alin As Long, nome As String)
    Dim er As Long, mt As Object
    If f1 Is Nothing Or f2 Is Nothing Then
        AddLog "AVISO: mate '" & nome & U("' n\u00E3o criado (refer\u00EAncia n\u00E3o encontrada).")
        Exit Sub
    End If
    md.ClearSelection2 True
    f1.Select2 False, 1
    f2.Select2 True, 1
    Passo "AddMate5: " & nome
    Set mt = asy.AddMate5(tipo, alin, False, 0, 0, 0, 1, 1, 0, 0, 0, False, False, 0, er)
    Passo "AddMate5: ok"
    If mt Is Nothing Then
        AddLog "AVISO: mate '" & nome & U("' falhou (c\u00F3digo ") & er & ")."
    Else
        AddLog "Mate criado: " & nome
    End If
    md.ClearSelection2 True
End Sub

'==============================================================================
' ANIMACAO CINEMATICA (montagem ativa)
'==============================================================================
Public Function PrepararAnimacao() As String
    Dim md As Object, cpm As Object, nR As String, nM As String
    If swApp Is Nothing Then Set swApp = Application.SldWorks
    Set md = swApp.ActiveDoc
    If md Is Nothing Then PrepararAnimacao = "Abra (ative) a montagem gerada pela macro.": Exit Function
    If md.GetType <> 2 Then PrepararAnimacao = U("O documento ativo n\u00E3o \u00E9 uma montagem."): Exit Function
    Set cpm = md.Extension.CustomPropertyManager("")
    nR = LerProp(cpm, "GEN_RODA")
    nM = LerProp(cpm, "GEN_MANIVELA")
    If nR = "" Or nM = "" Then PrepararAnimacao = U("Esta montagem n\u00E3o foi gerada pela macro (propriedades GEN_* ausentes)."): Exit Function
    CarregarCinematica CLng(Val(LerProp(cpm, "GEN_N"))), Val(LerProp(cpm, "GEN_C")), CLng(Val(LerProp(cpm, "GEN_PINOS")))
    Set A_md = md
    Set A_cR = md.GetComponentByName(nR)
    Set A_cM = md.GetComponentByName(nM)
    If A_cR Is Nothing Or A_cM Is Nothing Then PrepararAnimacao = "Componentes " & nR & " / " & nM & U(" n\u00E3o encontrados."): Exit Function
    If G_RotSign = 0 Then DetectarSinalRotacao
    PrepararAnimacao = ""
End Function

Private Function LerProp(cpm As Object, nome As String) As String
    Dim v As String, rv As String
    On Error Resume Next
    cpm.Get4 nome, False, v, rv
    If rv <> "" Then LerProp = rv Else LerProp = v
End Function

Private Sub AplicarPose(th As Double, reconstruir As Boolean)
    RotacionarZ A_cR, AnguloRoda(th)
    RotacionarZ A_cM, th
    If reconstruir Then A_md.EditRebuild3
    A_md.GraphicsRedraw2
End Sub

Public Function PosicionarTheta(th As Double, reconstruir As Boolean) As String
    Dim s As String
    s = PrepararAnimacao()
    If s <> "" Then PosicionarTheta = s: Exit Function
    AplicarPose th, reconstruir
End Function

Public Function Animar(voltas As Double, passos As Long, atrasoMs As Double, horario As Boolean, th0 As Double, reconstruir As Boolean) As String
    Dim s As String, i As Long, total As Long, sg As Double, th As Double
    s = PrepararAnimacao()
    If s <> "" Then Animar = s: Exit Function
    If passos < 4 Then passos = 4
    total = CLng(voltas * passos)
    If horario Then sg = -1 Else sg = 1
    G_PararAnim = False
    For i = 0 To total
        If G_PararAnim Then Exit For
        If GetAsyncKeyState(&H1B) <> 0 Then Exit For      ' ESC interrompe
        th = th0 + sg * i * 360# / passos
        AplicarPose th, reconstruir
        Esperar atrasoMs
    Next
    Animar = ""
End Function

Private Sub Esperar(ms As Double)
    Dim t0 As Single
    t0 = Timer
    Do While (Timer - t0) * 1000 < ms And Timer >= t0
        DoEvents
    Loop
    DoEvents
End Sub

'==============================================================================
' ARQUIVOS AUXILIARES
'==============================================================================
Private Sub GravarCSV(arq As String)
    Dim fnum As Integer, th As Double, passo As Double, h As Double
    Dim a0 As Double, a1 As Double, a2 As Double, d1 As Double, d2 As Double
    passo = PN("passoCSV"): If passo <= 0 Then passo = 1
    h = 0.01
    fnum = FreeFile
    Open arq For Output As #fnum
    Print #fnum, "theta_manivela_graus;angulo_roda_graus;w2_sobre_w1;alfa2_rad_s2"
    For th = 0 To 360.0000001 Step passo
        a0 = AnguloRoda(th - h): a1 = AnguloRoda(th): a2 = AnguloRoda(th + h)
        d1 = (a2 - a0) / (2 * h)
        d2 = (a2 - 2 * a1 + a0) / (h * h) * 180# / PI * G_w1 ^ 2
        Print #fnum, Format(th, "0.000") & ";" & Format(a1, "0.0000") & ";" & Format(d1, "0.00000") & ";" & Format(d2, "0.000")
    Next
    Close #fnum
    AddLog "CSV: " & arq
End Sub

Private Sub GravarMotion(arq As String, aA As String)
    Dim fnum As Integer, rpm As Double, T As Double
    rpm = PN("rpm"): If rpm <= 0 Then rpm = 60
    T = 3# * 60# / rpm
    fnum = FreeFile
    Open arq For Output As #fnum
    Print #fnum, "ESTUDO DE MOVIMENTO - MECANISMO GENEBRA"
    Print #fnum, "Montagem: " & aA
    Print #fnum, ""
    Print #fnum, U("Na montagem a base \u00E9 fixa e a roda e a manivela giram livres nos seus eixos.")
    Print #fnum, U("O sincronismo f\u00EDsico vem do CONTATO entre pino/disco de travamento e a roda.")
    Print #fnum, ""
    Print #fnum, "1. Ative o suplemento: Ferramentas > Suplementos > SOLIDWORKS Motion."
    Print #fnum, U("2. Clique na aba 'Estudo de movimento 1' (embaixo) e escolha o tipo 'An\u00E1lise de movimento'.")
    Print #fnum, U("   ('Movimento b\u00E1sico' tamb\u00E9m funciona, com menos precis\u00E3o.)")
    Print #fnum, U("3. Motor rotativo: selecione a face cil\u00EDndrica do furo ou do cubo da MANIVELA.")
    Print #fnum, "   Tipo: Velocidade constante = " & Format(rpm, "0.##") & " RPM. Confira o sentido pela seta."
    Print #fnum, U("4. Contato (corpos s\u00F3lidos): selecione os componentes RODA e MANIVELA.")
    Print #fnum, U("   Marque 'Material' (ex.: A\u00E7o seco) ou informe atrito; deixe 'Atrito' ligado.")
    Print #fnum, "5. Propriedades do estudo de movimento: Quadros por segundo = 200 (ou mais)."
    Print #fnum, U("   Se o pino 'atravessar' a ranhura, aumente a precis\u00E3o/qualidade do contato.")
    Print #fnum, U("6. Arraste a chave de tempo at\u00E9 ") & Format(T, "0.00") & " s (3 voltas da manivela)."
    Print #fnum, U("7. Calcular. Depois: Resultados e Plotagens > Deslocamento/Velocidade/Acelera\u00E7\u00E3o angular da RODA.")
    Print #fnum, ""
    Print #fnum, U("Valores esperados (c\u00E1lculo da macro):")
    Print #fnum, "  Movimento/parada: " & Format(G_frMov * 100, "0.0") & "% / " & Format((1 - G_frMov) * 100, "0.0") & "%"
    Print #fnum, U("  Tempo de indexa\u00E7\u00E3o: ") & Format(G_tIdx, "0.000") & " s   Parada: " & Format(G_tPar, "0.000") & " s"
    Print #fnum, U("  w2 m\u00E1x: ") & Format(G_w2max, "0.00") & U(" rad/s   Acelera\u00E7\u00E3o angular m\u00E1x.: ") & Format(G_a2max, "0.0") & U(" rad/s\u00B2")
    Close #fnum
End Sub

'==============================================================================
' UTILITARIOS DE DOCUMENTO / ESBOCO / FEATURES
'==============================================================================
Private Function Template(tipo As Long, campo As String) As String
    Dim t As String
    t = Trim(PS(campo))
    If t = "" Then t = swApp.GetUserPreferenceStringValue(tipo)
    If t = "" Then Err.Raise vbObjectError + 40, , U("Template padr\u00E3o n\u00E3o definido (Ferramentas > Op\u00E7\u00F5es > Modelos padr\u00E3o) - ou informe na aba Sa\u00EDda.")
    If Dir(t) = "" Then Err.Raise vbObjectError + 41, , U("Template n\u00E3o encontrado: ") & t
    Template = t
End Function

Private Function NovaPeca() As Object
    Dim md As Object
    Passo "NovaPeca: NewDocument"
    Set md = swApp.NewDocument(Template(8, "tplPeca"), 0, 0, 0)
    Passo "NovaPeca: ok"
    If md Is Nothing Then Err.Raise vbObjectError + 42, , U("N\u00E3o foi poss\u00EDvel criar pe\u00E7a (verifique o template).")
    If md.GetType <> 1 Then Err.Raise vbObjectError + 43, , U("O template de pe\u00E7a n\u00E3o \u00E9 um .prtdot.")
    Set NovaPeca = md
End Function

Private Sub FecharDoc(arq As String)
    Dim d As Object
    On Error Resume Next
    Set d = swApp.GetOpenDocumentByName(arq)
    If Not d Is Nothing Then swApp.CloseDoc d.GetTitle
End Sub

Private Sub Salvar(md As Object, arq As String)
    Dim e As Long, w As Long
    Passo "Propriedades Autor/Empresa"
    On Error Resume Next
    md.Extension.CustomPropertyManager("").Add3 "Autor", 30, "Gilmar Barbosa", 2
    md.Extension.CustomPropertyManager("").Add3 "Empresa", 30, "Autom", 2
    On Error GoTo 0
    Passo "SaveAs: " & arq
    If Not md.Extension.SaveAs(arq, 0, 1, Nothing, e, w) Then
        Err.Raise vbObjectError + 44, , "Falha ao salvar " & arq & " (erro " & e & ")."
    End If
    AddLog "Salvo: " & arq
End Sub

Private Sub CriarPasta(p As String)
    Dim pai As String, k As Long
    If Dir(p, vbDirectory) <> "" Then Exit Sub
    k = InStrRev(p, "\")
    If k > 3 Then
        pai = Left(p, k - 1)
        CriarPasta pai
    End If
    MkDir p
End Sub

' Devolve os 3 planos padrao (Frontal, Superior, Direito) independentemente do idioma
Private Function PlanosPadrao(md As Object) As Variant
    Dim f As Object, arr(2) As Object, k As Integer
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then
            Set arr(k) = f
            k = k + 1
            If k > 2 Then Exit Do
        End If
        Set f = f.GetNextFeature
    Loop
    If k < 3 Then Err.Raise vbObjectError + 45, , U("Planos padr\u00E3o n\u00E3o encontrados no template.")
    G_PlFront = arr(0).Name
    PlanosPadrao = arr
End Function

Private Sub IniciarEsboco(ByVal md As Object, ByVal plano As Object)
    Passo "### IniciarEsboco"
    md.ClearSelection2 True
    plano.Select2 False, 0
    md.SketchManager.InsertSketch True
    md.SketchManager.AddToDB = True
    md.SketchManager.DisplayWhenAdded = False
End Sub

Private Function FecharEsboco(md As Object) As Object
    Passo "### FecharEsboco"
    Dim sk As Object, f As SldWorks.Feature
    Set sk = md.SketchManager.ActiveSketch
    If sk Is Nothing Then Err.Raise vbObjectError + 46, , U("Nenhum esbo\u00E7o ativo.")
    Set f = sk
    md.SketchManager.AddToDB = False
    md.SketchManager.DisplayWhenAdded = True
    md.SketchManager.InsertSketch True
    Set FecharEsboco = f
End Function

Private Sub Extrudar(md As Object, sk As Object, d1 As Double, inverter As Boolean, nome As String, _
                     Optional doisLados As Boolean = False, Optional d2 As Double = 0)
    Dim f As Object
    md.ClearSelection2 True
    sk.Select2 False, 0
    Passo "Extrusao: " & nome
    Set f = md.FeatureManager.FeatureExtrusion2(Not doisLados, False, inverter, 0, 0, d1 * MM, d2 * MM, _
            False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    If f Is Nothing Then Err.Raise vbObjectError + 47, , U("Falha na extrus\u00E3o '") & nome & "'."
    Renomear f, nome
End Sub

Private Sub CortarPassante(md As Object, sk As Object, nome As String)
    Dim f As Object, tentativa As Integer
    For tentativa = 1 To 3
        md.ClearSelection2 True
        sk.Select2 False, 0
        Passo "Corte: " & nome & " tentativa " & tentativa
        Select Case tentativa
            Case 1  ' passante nos dois sentidos
                Set f = md.FeatureManager.FeatureCut4(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, _
                        0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
            Case 2
                Set f = md.FeatureManager.FeatureCut4(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, _
                        0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
            Case 3
                Set f = md.FeatureManager.FeatureCut4(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, _
                        0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        End Select
        If Not f Is Nothing Then Exit For
    Next
    If f Is Nothing Then Err.Raise vbObjectError + 48, , "Falha no corte '" & nome & "'."
    Renomear f, nome
End Sub

Private Sub EixoReferencia(ByVal md As Object, ByVal p1 As Object, ByVal p2 As Object, nome As String)
    md.ClearSelection2 True
    p1.Select2 False, 0
    p2.Select2 True, 0
    Passo "Eixo: " & nome
    If Not md.InsertAxis2(True) Then Err.Raise vbObjectError + 49, , "Falha ao criar o eixo " & nome
    Renomear md.FeatureByPositionReverse(0), nome
    md.ClearSelection2 True
End Sub

Private Function PlanoOffset(ByVal md As Object, ByVal base As Object, ByVal d As Double, ByVal idx As Integer, nome As String) As Object
    Dim f1 As Object, f2 As Object
    Set f1 = CriarPlano(md, base, d, 0)
    If PlanoConfere(f1, d, idx) Then
        Renomear f1, nome: Set PlanoOffset = f1: Exit Function
    End If
    Set f2 = CriarPlano(md, base, d, 256)      ' swRefPlaneReferenceConstraint_OptionFlip
    If PlanoConfere(f2, d, idx) Then
        Apagar md, f1
        Renomear f2, nome: Set PlanoOffset = f2: Exit Function
    End If
    Apagar md, f2
    If f1 Is Nothing Then Err.Raise vbObjectError + 50, , "Falha ao criar o plano " & nome
    AddLog U("AVISO: n\u00E3o foi poss\u00EDvel confirmar o lado do plano ") & nome & "; confira na base."
    Renomear f1, nome
    Set PlanoOffset = f1
End Function

Private Function CriarPlano(ByVal md As Object, ByVal base As Object, ByVal d As Double, ByVal flag As Long) As Object
    md.ClearSelection2 True
    base.Select2 False, 0
    Passo "Plano flag=" & flag
    Set CriarPlano = md.FeatureManager.InsertRefPlane(8 + flag, d * MM, 0, 0, 0, 0)
    md.ClearSelection2 True
End Function

Private Function PlanoConfere(f As Object, d As Double, idx As Integer) As Boolean
    Dim a As Variant
    If f Is Nothing Then Exit Function
    On Error Resume Next
    a = f.GetSpecificFeature2.Transform.ArrayData
    If Err.Number <> 0 Then Exit Function
    PlanoConfere = (Abs(a(9 + idx) - d * MM) < 0.000001)
End Function

Private Sub Apagar(md As Object, f As Object)
    If f Is Nothing Then Exit Sub
    md.ClearSelection2 True
    f.Select2 False, 0
    md.EditDelete
End Sub

Private Sub Renomear(f As Object, nome As String)
    On Error Resume Next
    If Not f Is Nothing Then f.Name = nome
End Sub

Private Sub AplicarMaterialCor(md As Object, mat As String, cor As String)
    Dim cfg As String, db As String, nomeMat As String, v As Variant
    On Error Resume Next
    cfg = md.ConfigurationManager.ActiveConfiguration.Name
    If Trim(mat) <> "" Then
        Passo "Material: " & mat
        md.SetMaterialPropertyName2 cfg, "SOLIDWORKS Materials", mat
        nomeMat = md.GetMaterialPropertyName2(cfg, db)
        If nomeMat = "" Then
            md.SetMaterialPropertyName2 cfg, "Materiais SOLIDWORKS", mat
            nomeMat = md.GetMaterialPropertyName2(cfg, db)
        End If
        If nomeMat = "" Then AddLog "AVISO: material '" & mat & U("' n\u00E3o encontrado na biblioteca (pe\u00E7a ") & md.GetTitle & ")."
    End If
    AplicarCor md, cor
End Sub

' Aplica a cor no nivel da peca e de cada corpo solido. O nivel do corpo fica
' acima da aparencia do material na hierarquia, entao a cor nao some quando o
' material (ex.: aco) traz aparencia propria. Array tipado em Double: arrays
' Variant em chamadas da API podem falhar em silencio.
Private Sub AplicarCor(ByVal md As SldWorks.ModelDoc2, ByVal cor As String)
    Dim c(8) As Double, v As Variant
    Dim pd As SldWorks.PartDoc, corpos As Variant, i As Long, bd As SldWorks.Body2
    Passo "Cor: " & cor
    If Not CorRGB(cor, c) Then
        If Trim(cor) <> "" And LCase(Trim(cor)) <> "(sem cor)" Then AddLog "AVISO: cor '" & cor & U("' n\u00E3o reconhecida (pe\u00E7a ") & md.GetTitle & ")."
        Exit Sub
    End If
    v = c
    On Error Resume Next
    md.Extension.SetMaterialPropertyValues v, 2, Empty         ' Sub (sem retorno); 2 = todas as configuracoes
    If Err.Number <> 0 Then
        Err.Clear
        md.MaterialPropertyValues = v
    End If
    Err.Clear
    Set pd = md
    If Not pd Is Nothing Then
        corpos = pd.GetBodies2(0, False)                       ' 0 = corpos solidos
        If IsArray(corpos) Then
            For i = LBound(corpos) To UBound(corpos)
                Set bd = corpos(i)
                bd.MaterialPropertyValues2(2, Empty) = v
            Next
        End If
    End If
    If Err.Number <> 0 Then AddLog "AVISO: falha ao aplicar a cor '" & cor & "' em " & md.GetTitle & ": " & Err.Description
    Err.Clear
    md.GraphicsRedraw2
    Passo "Cor: ok"
End Sub

' Cor dos eixos: campo corE; presets antigos sem o campo ficam sem cor.
Private Function CorEixo() As String
    CorEixo = PS("corE")
End Function

Private Function CorRGB(ByVal cor As String, c() As Double) As Boolean
    Dim r As Double, g As Double, b As Double
    Select Case LCase(Trim(cor))
        ' tons medios e saturados (tinta industrial); cores claras demais ficam lavadas
        Case "cinza": r = 0.55: g = 0.56: b = 0.58
        Case "cinza claro": r = 0.76: g = 0.77: b = 0.79
        Case "cinza escuro": r = 0.3: g = 0.31: b = 0.33
        Case "azul": r = 0.08: g = 0.3: b = 0.68
        Case "vermelho": r = 0.75: g = 0.08: b = 0.08
        Case "verde": r = 0.1: g = 0.52: b = 0.22
        Case "amarelo": r = 0.95: g = 0.74: b = 0.04
        Case "laranja": r = 0.95: g = 0.42: b = 0.04
        Case "preto": r = 0.08: g = 0.08: b = 0.09
        Case "branco": r = 0.92: g = 0.92: b = 0.92
        Case Else
            CorRGB = False: Exit Function
    End Select
    ' R, G, B, ambiente, difusa, especular, brilho, transparencia, emissao
    c(0) = r: c(1) = g: c(2) = b
    ' ambiente e difusa moderadas evitam o aspecto desbotado; especular da brilho de metal pintado
    c(3) = 0.45: c(4) = 0.8: c(5) = 0.5: c(6) = 0.4: c(7) = 0#: c(8) = 0#
    CorRGB = True
End Function

' ---- primitivas de esboco (entrada em mm) ----
Private Sub Circulo(sm As Object, xc As Double, yc As Double, raio As Double)
    sm.CreateCircle xc * MM, yc * MM, 0, (xc + raio) * MM, yc * MM, 0
End Sub

Private Sub Linha(sm As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    sm.CreateLine x1 * MM, y1 * MM, 0, x2 * MM, y2 * MM, 0
End Sub

' dirArco: +1 anti-horario, -1 horario (de P1 para P2)
Private Sub Arco(sm As Object, xc As Double, yc As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, dirArco As Integer)
    sm.CreateArc xc * MM, yc * MM, 0, x1 * MM, y1 * MM, 0, x2 * MM, y2 * MM, 0, dirArco
End Sub

Private Sub DesenharFuro(sm As Object, cx As Double, cy As Double, d As Double, comCh As Boolean, b As Double, t As Double, angDeg As Double)
    Dim rh As Double, yk As Double, a As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double
    rh = d / 2
    If (Not comCh) Or b <= 0 Or t <= 0 Or b >= d Then
        Call Circulo(sm, cx, cy, rh)
        Exit Sub
    End If
    yk = Sqr(rh ^ 2 - (b / 2) ^ 2)
    a = (angDeg - 90#) * PI / 180#
    Call RotP(b / 2, yk, a, x1, y1)
    Call RotP(-b / 2, yk, a, x2, y2)
    Call RotP(-b / 2, rh + t, a, x3, y3)
    Call RotP(b / 2, rh + t, a, x4, y4)
    Call Arco(sm, cx, cy, cx + x1, cy + y1, cx + x2, cy + y2, -1)
    Call Linha(sm, cx + x2, cy + y2, cx + x3, cy + y3)
    Call Linha(sm, cx + x3, cy + y3, cx + x4, cy + y4)
    Call Linha(sm, cx + x4, cy + y4, cx + x1, cy + y1)
End Sub

Private Sub RetArred(sm As Object, x0 As Double, y0 As Double, x1 As Double, y1 As Double, rc As Double)
    If rc <= 0 Then
        Call Linha(sm, x0, y0, x1, y0)
        Call Linha(sm, x1, y0, x1, y1)
        Call Linha(sm, x1, y1, x0, y1)
        Call Linha(sm, x0, y1, x0, y0)
    Else
        Call Linha(sm, x0 + rc, y0, x1 - rc, y0)
        Call Arco(sm, x1 - rc, y0 + rc, x1 - rc, y0, x1, y0 + rc, 1)
        Call Linha(sm, x1, y0 + rc, x1, y1 - rc)
        Call Arco(sm, x1 - rc, y1 - rc, x1, y1 - rc, x1 - rc, y1, 1)
        Call Linha(sm, x1 - rc, y1, x0 + rc, y1)
        Call Arco(sm, x0 + rc, y1 - rc, x0 + rc, y1, x0, y1 - rc, 1)
        Call Linha(sm, x0, y1 - rc, x0, y0 + rc)
        Call Arco(sm, x0 + rc, y0 + rc, x0, y0 + rc, x0 + rc, y0, 1)
    End If
End Sub

'==============================================================================
' TRANSFORMACOES (rotacao em Z mantendo a translacao do componente)
'==============================================================================
' A convencao do SOLIDWORKS e fixa (linhas da matriz = eixos do componente).
' O teste antigo com MathUtility/CreatePoint derrubava o SOLIDWORKS e foi removido.
Private Sub DetectarSinalRotacao()
    G_RotSign = 1
End Sub

Private Function AnguloZero(angDeg As Double) As Boolean
    Dim r As Double
    r = angDeg - 360# * Int(angDeg / 360#)
    AnguloZero = (r < 0.000001 Or r > 360# - 0.000001)
End Function

' Tipos explicitos (early binding): chamadas late-bound da MathUtility com
' arrays podem corromper a pilha do SOLIDWORKS.
Public Sub RotacionarZ(ByVal comp As SldWorks.Component2, ByVal angDeg As Double)
    Dim mu As SldWorks.MathUtility, xfAtual As SldWorks.MathTransform, xfNovo As SldWorks.MathTransform
    Dim a As Variant, t(15) As Double, v As Variant
    Dim cs As Double, sn As Double
    If G_RotSign = 0 Then DetectarSinalRotacao
    Passo "Rotacao: Transform2 (ler) " & comp.Name2
    Set xfAtual = comp.Transform2
    If xfAtual Is Nothing Then
        Passo "Rotacao: Transform2 vazio"
        Exit Sub
    End If
    a = xfAtual.ArrayData
    cs = Cos(angDeg * PI / 180#)
    sn = Sin(angDeg * PI / 180#) * G_RotSign
    t(0) = cs: t(1) = sn: t(2) = 0
    t(3) = -sn: t(4) = cs: t(5) = 0
    t(6) = 0: t(7) = 0: t(8) = 1
    t(9) = a(9): t(10) = a(10): t(11) = a(11)
    t(12) = 1
    v = t
    Passo "Rotacao: GetMathUtility"
    Set mu = swApp.GetMathUtility
    Passo "Rotacao: CreateTransform"
    Set xfNovo = mu.CreateTransform(v)
    Passo "Rotacao: Transform2 (gravar)"
    comp.Transform2 = xfNovo
    Passo "Rotacao: ok"
End Sub
