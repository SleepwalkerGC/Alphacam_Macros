'FORM================================================================================================================

Public EscolhaUsuario As Integer ' 0 para 0°, 90 para 90°, -1 para cancelar

Private Sub cmdZero_Click()
    EscolhaUsuario = 0
    Me.Hide
End Sub

Private Sub cmdNoventa_Click()
    EscolhaUsuario = 90
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    EscolhaUsuario = -1
    Me.Hide
End Sub

Private Sub UserForm_Click()

End Sub

'End FORM ==========================================================================================================

Public Function InitAlphacamAddIn(acamversion As Long) As Integer
    Dim fr As Frame
    Set fr = App.Frame
    With fr
        ' set up itemname  and menuname as new string variables
        Dim ItemName As String, MenuName As String
        ItemName = "Alinhar Solido": MenuName = "VBA Tab"
        
        ' create the new menu
        .AddMenuItem2 ItemName, "AlinharSolido", acamMenuNEW, MenuName
        fr.AddButton fr.CreateButtonBar("VBA Tab"), "Alinhar_Solido2.bmp", fr.LastMenuCommandID
        
    End With
    InitAlphacamAddIn = 0
    
End Function


Sub AlinharSolido()
    Dim Drw As Drawing
    Set Drw = App.ActiveDrawing
    Drw.Font = "TArial"
    
    Dim Vista As ViewWindow
    Set Vista = App.ActiveDrawing.CurrentViewWindow
    Vista.View = acamViewTOP

    Dim P1 As Path, P2 As Path
    
    ' Selecionar a linha de referências
    Set P1 = Drw.UserSelectOneGeo("ANGULOS: SELECIONE UMA LINHA DE REFERÊNCIA")

    If P1 Is Nothing Then
        MsgBox "Nenhuma linha foi selecionada!", vbExclamation, "Erro"
        Exit Sub
    End If

    ' Calcular o ângulo da linha
    Dim E As Element
    Dim DeltaX As Double, DeltaY As Double, Ang As Double

    Set E = P1.Elements(1) ' Pegar o primeiro elemento da linha
    DeltaX = E.EndXL - E.StartXL
    DeltaY = E.EndYL - E.StartYL

    Ang = Atn(DeltaY / DeltaX) * 180 / 3.14159265358979
    
    Dim frm As New UserForm
    frm.Show vbModal

    If frm.EscolhaUsuario = -1 Then Exit Sub ' Cancelar operação

    Dim AnguloAlvo As Double
    If frm.EscolhaUsuario = 0 Then
    AnguloAlvo = -Ang  ' Alinhar para 0°
    End If
    
    If frm.EscolhaUsuario = 90 Then
    AnguloAlvo = -Ang + 90 ' Alinhar para 90°
    End If

    ' Aplicar rotação ao sólido
    Dim Solido As SolidFeatures
    Set Solido = App.ActiveDrawing.SolidInterface
    Dim Selector As solidselector
    Set Selector = Solido.Selector
    Selector.What = FeatureSelectBody
    Selector.Single = True
    Selector.Select "Selecione o sólido para girar"
    
    If Selector.Count = 1 Then
    Dim sBody As SolidBody
    Set sBody = Selector.Item(1)

    End If
    If Solido Is Nothing Then
        MsgBox "Nenhum sólido foi selecionado!", vbExclamation, "Erro"
        Exit Sub
    End If

    ' Obter ponto central da peça para rotação
    Dim CentroX As Double, CentroY As Double, CentroZ As Double
    CentroX = (sBody.MaxX + sBody.MinX) / 2
    CentroY = (sBody.MaxY + sBody.MinY) / 2
    CentroZ = (sBody.MaxZ + sBody.MinZ) / 2
    
    ' Executar rotação
    sBody.RotateG AnguloAlvo, CentroX, CentroY, CentroZ, CentroX, CentroY, CentroZ + 1
    
    Dim X1 As Double, Y1 As Double, Z1 As Double, X2 As Double, Y2 As Double, Z2 As Double, MoveX As Double, MoveY As Double
    
    MoveX = sBody.MinX
    MoveY = sBody.MinY
    
    sBody.MoveG -MoveX, -MoveY, 0
    
    X1 = sBody.MinX
    X2 = sBody.MaxX
    Y1 = sBody.MinY
    Y2 = sBody.MaxY
    Z1 = sBody.MinZ
    Z2 = sBody.MaxZ
    
    Set P2 = Drw.CreateRectangle(X1, Y1, X2, Y2)
    P2.SetMaterial Z1, Z2

    ' Atualizar exibição
    Drw.Clear True, False, False, True, False, False, False, False  'drw.Clear (Geometry, Construction, Toolpaths, Dimensions, Splines, Surfaces, UserLayers)
    Drw.Refresh
    Drw.RedrawShadedViews
    
End Sub