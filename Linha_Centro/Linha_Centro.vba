Public Function InitAlphacamAddIn(acamversion As Long) As Integer
    Dim fr As Frame
    Set fr = App.Frame
    With fr
        ' set up itemname  and menuname as new string variables
        Dim ItemName As String, MenuName As String
        ItemName = "Linha Centro": MenuName = "VBA Tab"
        
        ' create the new menu
        .AddMenuItem2 ItemName, "Linha_Centro", acamMenuNEW, MenuName
        fr.AddButton fr.CreateButtonBar("VBA Tab"), "Linha_Centro.bmp", fr.LastMenuCommandID
        
    End With
    InitAlphacamAddIn = 0
    
End Function

Sub Linha_Centro()

    Dim Drw As Drawing, P1 As Path, MinX As Double, MaxX As Double, MinY As Double, MaxY As Double, ZBottom As Double, _
        ZTop As Double, MidX As Double, MidY As Double, Lyr As Layer
    
    Set Drw = App.ActiveDrawing
    
    Set P1 = Drw.UserSelectOneGeo("SELECIONE UMA GEOMETRIA")
    
    If P1 Is Nothing Then
        MsgBox "Nenhuma geometria válida selecionada.", vbExclamation
        Exit Sub
    End If
    
    MinX = P1.MinXL
    MaxX = P1.MaxXL
    MinY = P1.MinYL
    MaxY = P1.MaxYL
    
    MidX = ((MaxX - MinX) / 2) + MinX
    MidY = ((MaxY - MinY) / 2) + MinY

    ZBottom = P1.Attribute("LicomUKDMBGeoZLevelBottom")
    ZTop = P1.Attribute("LicomUKDMBGeoZLevelTop")
    
    'ZBottom = CDbl(P1.Attribute("LicomUKDMBGeoZLevelBottom")) --> força valor numerico
    'ZTop = CDbl(P1.Attribute("LicomUKDMBGeoZLevelTop")) --> força valor numerico
    
    Set Lyr = Drw.CreateLayer("SLOT_GAV")
         Lyr.Color = acamCYAN
         Drw.SetLayer Lyr
    
    If MaxX > MaxY Then
    Drw.Create2DLine MinX, MidY, MaxX, MidY
    End If
    
    If MaxX < MaxY Then
    Drw.Create2DLine MidX, MinY, MidX, MaxY
    End If
    
    Drw.GetLastGeo.Attribute("LicomUKDMBGeoZLevelTop") = ZTop#
    Drw.GetLastGeo.Attribute("LicomUKDMBGeoZLevelBottom") = ZBottom#
    
    Drw.SetLayer Nothing
    App.ActiveDrawing.RedrawShadedViews
    
    ' Exibe os valores em uma mensagem formatada
'    MsgBox "Medidas:" & vbNewLine & _
'           "MinX: " & MinX & vbNewLine & _
'           "MaxX: " & MaxX & vbNewLine & _
'           "MinY: " & MinY & vbNewLine & _
'           "MaxY: " & MaxY & vbNewLine & _
'           "Z Bottom: " & ZBottom & vbNewLine & _
'           "Z Top: " & ZTop & vbNewLine


End Sub