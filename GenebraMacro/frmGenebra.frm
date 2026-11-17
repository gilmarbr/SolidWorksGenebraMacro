VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenebra 
   Caption         =   "Genebra"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmGenebra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'==============================================================================
' C?DIGO DO USERFORM "frmGenebra"
' Crie um UserForm vazio, mude o (Name) para frmGenebra, abra o c?digo (F7),
' apague tudo e cole ESTE texto inteiro. Os controles s?o criados por c?digo.
'==============================================================================
Option Explicit

Private mp As MSForms.MultiPage
Private Campos As Object          ' id -> controle
Private Tipos As Object           ' id -> "N" n?mero, "T" texto, "C" combo, "K" check
Private Eventos As Collection
Private txtRes As MSForms.TextBox
Private lblRef As MSForms.Label
Private lblTheta As MSForms.Label
Private scrTheta As MSForms.ScrollBar
Private Iniciando As Boolean

Private Const LBLW As Single = 150
Private Const CTLW As Single = 80
Private Const ROWH As Single = 21
Private Const COLW As Single = 256

'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim y As Single
    Iniciando = True
    If swApp Is Nothing Then Set swApp = Application.SldWorks
    Set Campos = CreateObject("Scripting.Dictionary")
    Set Tipos = CreateObject("Scripting.Dictionary")
    Set Eventos = New Collection

    Me.Caption = "Gerador de Mecanismo Genebra - SOLIDWORKS 2024"
    Me.Width = 540
    Me.Height = 525

    Set mp = Me.Controls.Add("Forms.MultiPage.1", "mpAbas")
    mp.Left = 6: mp.Top = 6: mp.Width = 522: mp.Height = 440
    Do While mp.Pages.Count < 7
        mp.Pages.Add
    Loop
    mp.Pages(0).Caption = "Geral"
    mp.Pages(1).Caption = "Roda"
    mp.Pages(2).Caption = "Manivela"
    mp.Pages(3).Caption = "Base e eixos"
    mp.Pages(4).Caption = "Sa?da"
    mp.Pages(5).Caption = "Resultados"
    mp.Pages(6).Caption = "Movimento"

    MontarGeral mp.Pages(0)
    MontarRoda mp.Pages(1)
    MontarManivela mp.Pages(2)
    MontarBase mp.Pages(3)
    MontarSaida mp.Pages(4)
    MontarResultados mp.Pages(5)
    MontarMovimento mp.Pages(6)

    y = 452
    AddBotao Me, "calc", "Calcular", 6, y, 80
    AddBotao Me, "salvar", "Salvar predef.", 92, y, 95
    AddBotao Me, "carregar", "Carregar predef.", 193, y, 100
    AddBotao Me, "gerar", "GERAR", 340, y, 90
    AddBotao Me, "fechar", "Fechar", 436, y, 90

    mp.Value = 0
    Iniciando = False
    AtualizarRotuloRef
    Recalcular
End Sub

Private Sub UserForm_Terminate()
    Set Eventos = Nothing
End Sub

'------------------------------------------------------------------------------
' ABAS
'------------------------------------------------------------------------------
Private Sub MontarGeral(pg As Object)
    AddNum pg, "n", "N? de ranhuras (3 a 18)", 6, 0, 0
    AddCombo pg, "pinos", "N? de pinos na manivela", "1|2", 0, 1, 0, False
    AddCombo pg, "modo", "Modo de dimensionamento", "Dist?ncia entre centros (C)|Raio da manivela (r)|Di?metro externo da roda", 0, 0, 1, False, 250
    AddNum pg, "ref", "Valor de refer?ncia", 100, 0, 2, "mm"
    Set lblRef = pg.Controls("lbl_ref")
    AddNum pg, "gap", "Folga axial entre pe?as", 0.5, 1, 2, "mm"
    AddCombo pg, "sentido", "Sentido da manivela", "Anti-hor?rio|Hor?rio", 0, 0, 3, False
    AddNum pg, "rpm", "Rota??o da manivela", 60, 1, 3, "rpm"
    AddNum pg, "theta0", "?ngulo inicial da manivela", 0, 0, 4, "?"
    AddInfo pg, 0, 6, 500, 120, _
        "Conven??es:" & vbCrLf & _
        "- Centro da roda na origem da montagem; centro da manivela em (C, 0)." & vbCrLf & _
        "- ?ngulo da manivela medido a partir de +X, anti-hor?rio visto de +Z." & vbCrLf & _
        "- 0? = pino apontando para fora da roda (meio da parada)." & vbCrLf & _
        "- Com 2 pinos, 90? come?a em parada." & vbCrLf & _
        "- Campos com '0 = auto' s?o calculados pela macro (veja a aba Resultados)."
End Sub

Private Sub MontarRoda(pg As Object)
    AddNum pg, "tR", "Espessura da roda", 10, 0, 0, "mm"
    AddNum pg, "fs", "Folga pino/ranhura", 0.2, 1, 0, "mm"
    AddNum pg, "ffundo", "Folga no fundo da ranhura", 1, 0, 1, "mm"
    AddNum pg, "RoMan", "Raio externo (0 = auto)", 0, 1, 1, "mm"
    AddNum pg, "fl", "Folga do arco de travamento", 0.2, 0, 2, "mm"
    AddNum pg, "ePonta", "Espessura da ponta do garfo", 1.5, 1, 2, "mm"
    AddNum pg, "dFuroR", "? do furo (eixo)", 12, 0, 3, "mm"
    AddChk pg, "chavR", "Rasgo de chaveta", False, 1, 3
    AddNum pg, "bChR", "Largura da chaveta", 4, 0, 4, "mm"
    AddNum pg, "tChR", "Profundidade da chaveta", 2.5, 1, 4, "mm"
    AddChk pg, "cuboR", "Cubo (lado da frente)", True, 0, 5
    AddNum pg, "dCuboR", "? do cubo", 30, 1, 5, "mm"
    AddNum pg, "hCuboR", "Altura do cubo", 8, 0, 6, "mm"
    AddCombo pg, "matR", "Material (nome da biblioteca)", ListaMateriais(), "AISI 1020", 0, 7, True, 250
    AddCombo pg, "corR", "Cor", ListaCores(), "(sem cor)", 0, 8, False
End Sub

Private Sub MontarManivela(pg As Object)
    AddNum pg, "dp", "? do pino", 10, 0, 0, "mm"
    AddNum pg, "hp", "Altura do pino (0 = auto)", 0, 1, 0, "mm"
    AddNum pg, "tM", "Espessura da placa", 8, 0, 1, "mm"
    AddCombo pg, "forma", "Formato da placa", "Disco|Bra?o", 0, 1, 1, False
    AddNum pg, "bBraco", "Largura do bra?o (0 = auto)", 0, 0, 2, "mm"
    AddNum pg, "RmMan", "Raio do disco (0 = auto)", 0, 1, 2, "mm"
    AddNum pg, "tL", "Esp. disco travamento (0=auto)", 0, 0, 3, "mm"
    AddNum pg, "RsMan", "Raio Rs travamento (0=auto)", 0, 1, 3, "mm"
    AddNum pg, "fa", "Folga do al?vio", 0.5, 0, 4, "mm"
    AddNum pg, "dFuroM", "? do furo (eixo)", 12, 1, 4, "mm"
    AddChk pg, "chavM", "Rasgo de chaveta", False, 0, 5
    AddNum pg, "bChM", "Largura da chaveta", 4, 1, 5, "mm"
    AddNum pg, "tChM", "Profundidade da chaveta", 2.5, 0, 6, "mm"
    AddChk pg, "cuboM", "Cubo (lado de tr?s)", True, 1, 6
    AddNum pg, "dCuboM", "? do cubo", 30, 0, 7, "mm"
    AddNum pg, "hCuboM", "Altura do cubo", 8, 1, 7, "mm"
    AddCombo pg, "matM", "Material (nome da biblioteca)", ListaMateriais(), "AISI 1020", 0, 8, True, 250
    AddCombo pg, "corM", "Cor", ListaCores(), "(sem cor)", 0, 9, False
End Sub

Private Sub MontarBase(pg As Object)
    AddChk pg, "gBase", "Gerar placa base (corpo)", True, 0, 0
    AddNum pg, "tB", "Espessura da base", 8, 1, 0, "mm"
    AddNum pg, "mB", "Margem em volta", 15, 0, 1, "mm"
    AddNum pg, "rcB", "Raio dos cantos", 5, 1, 1, "mm"
    AddCombo pg, "nFix", "Furos de fixa??o", "0|4", 1, 0, 2, False
    AddNum pg, "dFix", "? furos de fixa??o", 6.5, 1, 2, "mm"
    AddCombo pg, "matB", "Material da base", ListaMateriais(), "6061-T6 (SS)", 0, 3, True, 250
    AddCombo pg, "corB", "Cor da base", ListaCores(), "(sem cor)", 0, 4, False
    AddChk pg, "gEixos", "Gerar eixos (pe?as)", True, 0, 5
    AddNum pg, "sobra", "Sobra do eixo acima da pe?a", 2, 1, 5, "mm"
    AddCombo pg, "matE", "Material dos eixos", ListaMateriais(), "AISI 1045 A?o, estirado a frio", 0, 6, True, 250
    AddCombo pg, "corE", "Cor dos eixos", ListaCores(), "(sem cor)", 0, 7, False
    AddInfo pg, 0, 8, 500, 80, _
        "Os eixos usam o mesmo ? dos furos da roda e da manivela." & vbCrLf & _
        "Sem placa base, a pe?a 'Base' ? criada s? com eixos e planos de refer?ncia," & vbCrLf & _
        "que a montagem usa nos mates."
End Sub

Private Sub MontarSaida(pg As Object)
    AddTxt pg, "pasta", "Pasta de sa?da", Environ("USERPROFILE") & "\Documents\Genebra", 0, 0, 290
    AddBotao pg, "procurar", "Procurar...", 452, 10 + 0 * ROWH - 1, 60
    AddTxt pg, "prefixo", "Prefixo (vazio = auto)", "", 0, 1, 150
    AddChk pg, "sobrescrever", "Sobrescrever arquivos", False, 0, 2
    AddChk pg, "gCSV", "Gerar CSV da cinem?tica", True, 1, 2
    AddNum pg, "passoCSV", "Passo do CSV", 1, 0, 3, "?"
    AddTxt pg, "tplPeca", "Template de pe?a (opcional)", "", 0, 4, 340
    AddTxt pg, "tplMont", "Template de montagem (opcional)", "", 0, 5, 340
    AddInfo pg, 0, 7, 500, 90, _
        "Prefixo autom?tico: GEN_n<ranhuras>_C<dist?ncia>_" & vbCrLf & _
        "Arquivos: Roda, Manivela, Base, EixoRoda, EixoManivela, Montagem," & vbCrLf & _
        "Cinematica.csv e LEIA-ME_Motion.txt." & vbCrLf & _
        "Templates em branco = modelos padr?o de Ferramentas > Op??es." & vbCrLf & _
        "A predefini??o ? salva como GenebraPreset.ini na pasta de sa?da."
End Sub

Private Sub MontarResultados(pg As Object)
    Set txtRes = pg.Controls.Add("Forms.TextBox.1", "txtRes")
    With txtRes
        .Left = 6: .Top = 6: .Width = 504: .Height = 400
        .MultiLine = True
        .ScrollBars = 2
        .WordWrap = False
        .Locked = True
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
End Sub

Private Sub MontarMovimento(pg As Object)
    Dim lb As MSForms.Label, ev As clsEvento
    Set lb = pg.Controls.Add("Forms.Label.1")
    lb.Caption = "?ngulo da manivela": lb.Left = 8: lb.Top = 14: lb.Width = 110
    Set scrTheta = pg.Controls.Add("Forms.ScrollBar.1", "scrTheta")
    With scrTheta
        .Left = 120: .Top = 10: .Width = 300: .Height = 18
        .Min = 0: .Max = 360: .SmallChange = 1: .LargeChange = 15
        .Orientation = 1
    End With
    Set ev = New clsEvento
    Set ev.Owner = Me: ev.Tag = "theta": Set ev.Scr = scrTheta
    Eventos.Add ev
    Set lblTheta = pg.Controls.Add("Forms.Label.1")
    lblTheta.Left = 428: lblTheta.Top = 14: lblTheta.Width = 70: lblTheta.Caption = "0?"

    AddChk pg, "aoVivo", "Atualizar ao arrastar", True, 0, 2
    AddChk pg, "rebuild", "Reconstruir a cada passo", False, 1, 2
    AddNum pg, "voltas", "Voltas da manivela", 2, 0, 3
    AddNum pg, "passos", "Passos por volta", 72, 1, 3
    AddNum pg, "atraso", "Atraso entre passos", 20, 0, 4, "ms"
    AddBotao pg, "posicionar", "Posicionar", 8, 10 + 5 * ROWH + 4, 90
    AddBotao pg, "animar", "Animar", 104, 10 + 5 * ROWH + 4, 90
    AddBotao pg, "parar", "Parar", 200, 10 + 5 * ROWH + 4, 90
    AddInfo pg, 0, 7, 500, 120, _
        "Abra e ATIVE a montagem gerada (ela guarda os dados em propriedades GEN_*)." & vbCrLf & _
        "A macro gira roda e manivela pela cinem?tica exata do Genebra." & vbCrLf & _
        "O sentido de rota??o usa o campo da aba Geral." & vbCrLf & _
        "Se a tela n?o atualizar, marque 'Reconstruir a cada passo'." & vbCrLf & _
        "Para simula??o f?sica (contato + motor) siga o arquivo LEIA-ME_Motion.txt."
End Sub

'------------------------------------------------------------------------------
' CRIA??O DE CONTROLES
'------------------------------------------------------------------------------
Private Function XCol(col As Integer) As Single
    XCol = 8 + col * COLW
End Function

Private Function YRow(row As Integer) As Single
    YRow = 10 + row * ROWH
End Function

Private Sub AddRotulo(pg As Object, id As String, cap As String, col As Integer, row As Integer)
    Dim lb As MSForms.Label
    Set lb = pg.Controls.Add("Forms.Label.1", "lbl_" & id)
    lb.Caption = cap
    lb.Left = XCol(col): lb.Top = YRow(row) + 3
    lb.Width = LBLW: lb.Height = 14
End Sub

Private Sub AddNum(pg As Object, id As String, cap As String, padrao As Double, col As Integer, row As Integer, Optional unid As String = "")
    Dim tB As MSForms.TextBox
    If unid <> "" Then cap = cap & " [" & unid & "]"
    AddRotulo pg, id, cap, col, row
    Set tB = pg.Controls.Add("Forms.TextBox.1", "f_" & id)
    tB.Left = XCol(col) + LBLW + 2: tB.Top = YRow(row)
    tB.Width = CTLW: tB.Height = 17
    tB.Text = CStr(padrao)
    Registrar id, "N", tB
End Sub

Private Sub AddTxt(pg As Object, id As String, cap As String, padrao As String, col As Integer, row As Integer, largura As Single)
    Dim tB As MSForms.TextBox
    AddRotulo pg, id, cap, col, row
    Set tB = pg.Controls.Add("Forms.TextBox.1", "f_" & id)
    tB.Left = XCol(col) + LBLW + 2: tB.Top = YRow(row)
    tB.Width = largura: tB.Height = 17
    tB.Text = padrao
    Registrar id, "T", tB
End Sub

' itens separados por "|". padrao: ?ndice (n?mero) ou texto.
Private Sub AddCombo(pg As Object, id As String, cap As String, itens As String, padrao As Variant, col As Integer, row As Integer, editavel As Boolean, Optional largura As Single = 0)
    Dim cB As MSForms.ComboBox, v As Variant, i As Long
    AddRotulo pg, id, cap, col, row
    Set cB = pg.Controls.Add("Forms.ComboBox.1", "f_" & id)
    cB.Left = XCol(col) + LBLW + 2: cB.Top = YRow(row)
    If largura > 0 Then cB.Width = largura Else cB.Width = CTLW + 20
    cB.Height = 17
    v = Split(itens, "|")
    For i = LBound(v) To UBound(v)
        cB.AddItem v(i)
    Next
    If editavel Then cB.Style = 0 Else cB.Style = 2
    If VarType(padrao) = vbString Then
        ' seleciona pelo indice: .Text em lista suspensa (Style 2) pode nao pegar
        For i = 0 To cB.ListCount - 1
            If LCase(cB.List(i)) = LCase(CStr(padrao)) Then cB.ListIndex = i: Exit For
        Next
        If cB.ListIndex < 0 And editavel Then cB.Text = padrao
    Else
        cB.ListIndex = CLng(padrao)
    End If
    Registrar id, "C", cB
End Sub

Private Sub AddChk(pg As Object, id As String, cap As String, padrao As Boolean, col As Integer, row As Integer)
    Dim ck As MSForms.CheckBox
    Set ck = pg.Controls.Add("Forms.CheckBox.1", "f_" & id)
    ck.Caption = cap
    ck.Left = XCol(col): ck.Top = YRow(row)
    ck.Width = LBLW + CTLW: ck.Height = 17
    ck.Value = padrao
    Registrar id, "K", ck
End Sub

Private Sub AddInfo(pg As Object, col As Integer, row As Integer, largura As Single, altura As Single, txt As String)
    Dim lb As MSForms.Label
    Set lb = pg.Controls.Add("Forms.Label.1")
    lb.Left = XCol(col): lb.Top = YRow(row)
    lb.Width = largura: lb.Height = altura
    lb.WordWrap = True
    lb.ForeColor = RGB(70, 70, 70)
    lb.Caption = txt
End Sub

Private Sub AddBotao(parent As Object, id As String, cap As String, x As Single, y As Single, largura As Single)
    Dim bt As MSForms.CommandButton, ev As clsEvento
    Set bt = parent.Controls.Add("Forms.CommandButton.1", "b_" & id)
    bt.Caption = cap
    bt.Left = x: bt.Top = y: bt.Width = largura: bt.Height = 22
    Set ev = New clsEvento
    Set ev.Owner = Me: ev.Tag = id: Set ev.Btn = bt
    Eventos.Add ev
End Sub

Private Sub Registrar(id As String, kind As String, ctl As Object)
    Dim ev As clsEvento
    Campos.Add id, ctl
    Tipos.Add id, kind
    Set ev = New clsEvento
    Set ev.Owner = Me
    ev.Tag = id
    Select Case kind
        Case "N", "T": Set ev.txt = ctl
        Case "C": Set ev.Cmb = ctl
        Case "K": Set ev.Chk = ctl
    End Select
    Eventos.Add ev
End Sub

Private Function ListaMateriais() As String
    ListaMateriais = "AISI 1020|AISI 1045 A?o, estirado a frio|A?o carbono comum|6061-T6 (SS)|Liga 6061-T6 (SS)|7075-T6 (SN)|Acetal (POM)|Nylon 6/10|Lat?o"
End Function

Private Function ListaCores() As String
    ListaCores = "(sem cor)|Cinza|Cinza claro|Cinza escuro|Azul|Vermelho|Verde|Amarelo|Laranja|Preto|Branco"
End Function

'------------------------------------------------------------------------------
' EVENTOS (chamados por clsEvento)
'------------------------------------------------------------------------------
Public Sub OnCampo(id As String)
    If Iniciando Then Exit Sub
    If id = "modo" Then AtualizarRotuloRef
    If id = "aoVivo" Or id = "rebuild" Or id = "voltas" Or id = "passos" Or id = "atraso" Then Exit Sub
    Recalcular
End Sub

Public Sub OnScroll(id As String)
    Dim msg As String
    If Iniciando Then Exit Sub
    lblTheta.Caption = scrTheta.Value & "?"
    If Campos("aoVivo").Value Then
        msg = PosicionarTheta(CDbl(scrTheta.Value), CBool(Campos("rebuild").Value))
        If msg <> "" Then lblTheta.Caption = scrTheta.Value & "? (!)"
    End If
End Sub

Public Sub OnBotao(id As String)
    Dim msg As String
    Select Case id
        Case "calc"
            Recalcular
            mp.Value = 5
        Case "gerar"
            LerParametros
            txtRes.Text = "Gerando... aguarde."
            mp.Value = 5
            DoEvents
            txtRes.Text = Gerar()
        Case "salvar"
            SalvarPreset
        Case "carregar"
            CarregarPreset
        Case "procurar"
            ProcurarPasta
        Case "fechar"
            G_PararAnim = True
            Unload Me
        Case "posicionar"
            msg = PosicionarTheta(CDbl(scrTheta.Value), CBool(Campos("rebuild").Value))
            If msg <> "" Then MsgBox msg, vbExclamation
        Case "animar"
            LerParametros
            msg = Animar(PN("voltas"), CLng(PN("passos")), PN("atraso"), _
                         (CLng(PN("sentido")) = 1), CDbl(scrTheta.Value), PB("rebuild"))
            If msg <> "" Then MsgBox msg, vbExclamation
        Case "parar"
            G_PararAnim = True
    End Select
End Sub

'------------------------------------------------------------------------------
' PAR?METROS / C?LCULO
'------------------------------------------------------------------------------
Private Function ParseNum(s As String) As Double
    ParseNum = Val(Replace(Trim(s), ",", "."))
End Function

Private Sub LerParametros()
    Dim k As Variant, C As Object
    Set Prm = CreateObject("Scripting.Dictionary")
    For Each k In Campos.Keys
        Set C = Campos(k)
        Select Case Tipos(k)
            Case "N"
                Prm(k) = ParseNum(C.Text)
            Case "T"
                Prm(k & "$") = C.Text
            Case "C"
                Prm(k) = C.ListIndex
                Prm(k & "$") = C.Text
            Case "K"
                If C.Value Then Prm(k) = 1 Else Prm(k) = 0
        End Select
    Next
    ' combos num?ricos: usam o texto do item
    Prm("pinos") = Val(Prm("pinos$"))
    Prm("nFix") = Val(Prm("nFix$"))
End Sub

Private Sub Recalcular()
    Dim msgs As New Collection, ok As Boolean
    On Error GoTo Falha
    LerParametros
    ok = Calcular(msgs)
    txtRes.Text = TextoResultados(ok, msgs)
    Exit Sub
Falha:
    txtRes.Text = "Valores inv?lidos: " & Err.Description
End Sub

Private Sub AtualizarRotuloRef()
    If lblRef Is Nothing Then Exit Sub
    Select Case Campos("modo").ListIndex
        Case 1: lblRef.Caption = "Raio da manivela r [mm]"
        Case 2: lblRef.Caption = "Di?metro externo da roda [mm]"
        Case Else: lblRef.Caption = "Dist?ncia entre centros C [mm]"
    End Select
End Sub

'------------------------------------------------------------------------------
' PREDEFINI??ES
'------------------------------------------------------------------------------
Private Function ArqPreset() As String
    Dim p As String
    p = Trim(Campos("pasta").Text)
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    ArqPreset = p & "\GenebraPreset.ini"
End Function

Private Sub SalvarPreset()
    Dim fnum As Integer, k As Variant, v As String, p As String
    On Error GoTo Falha
    p = Trim(Campos("pasta").Text)
    If p = "" Then MsgBox "Informe a pasta de sa?da.", vbExclamation: Exit Sub
    If Dir(p, vbDirectory) = "" Then MkDir p
    fnum = FreeFile
    Open ArqPreset() For Output As #fnum
    For Each k In Campos.Keys
        If Tipos(k) = "K" Then
            If Campos(k).Value Then v = "1" Else v = "0"
        Else
            v = Campos(k).Text
        End If
        Print #fnum, k & "=" & v
    Next
    Close #fnum
    MsgBox "Predefini??o salva em:" & vbCrLf & ArqPreset(), vbInformation
    Exit Sub
Falha:
    MsgBox "N?o foi poss?vel salvar: " & Err.Description, vbExclamation
End Sub

Private Sub CarregarPreset()
    Dim fnum As Integer, Linha As String, k As String, v As String, p As Long
    On Error GoTo Falha
    If Dir(ArqPreset()) = "" Then MsgBox "Arquivo n?o encontrado:" & vbCrLf & ArqPreset(), vbExclamation: Exit Sub
    Iniciando = True
    fnum = FreeFile
    Open ArqPreset() For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, Linha
        p = InStr(Linha, "=")
        If p > 0 Then
            k = Left(Linha, p - 1): v = Mid(Linha, p + 1)
            If Campos.Exists(k) Then
                If Tipos(k) = "K" Then
                    Campos(k).Value = (v = "1")
                ElseIf Tipos(k) = "C" Then
                    SelecionarCombo Campos(k), v
                Else
                    Campos(k).Text = v
                End If
            End If
        End If
    Loop
    Close #fnum
    Iniciando = False
    AtualizarRotuloRef
    Recalcular
    Exit Sub
Falha:
    Iniciando = False
    MsgBox "N?o foi poss?vel carregar: " & Err.Description, vbExclamation
End Sub

' Lista suspensa (Style 2): escolhe pelo indice; combo editavel aceita texto livre.
Private Sub SelecionarCombo(cb As Object, v As String)
    Dim i As Long
    For i = 0 To cb.ListCount - 1
        If LCase(cb.List(i)) = LCase(v) Then cb.ListIndex = i: Exit Sub
    Next
    If cb.Style = 0 Then cb.Text = v
End Sub

Private Sub ProcurarPasta()
    Dim sh As Object, fo As Object
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    Set fo = sh.BrowseForFolder(0, "Pasta de sa?da", 0)
    If Not fo Is Nothing Then Campos("pasta").Text = fo.Self.Path
End Sub


