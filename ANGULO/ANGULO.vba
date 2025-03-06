Public Function InitAlphacamAddIn(acamversion As Long) As Integer
    Dim fr As Frame
    Set fr = App.Frame
    With fr
        ' set up itemname  and menuname as new string variables
        Dim ItemName As String, MenuName As String, ItemName2 As String
        ItemName = "RELATIVO": MenuName = "VBA Tab": ItemName2 = "ABSOLUTO"
        
        ' create the new menu
        .AddMenuItem2 ItemName, "Angulo_Relativo", acamMenuNEW, MenuName
        .AddMenuItem2 ItemName2, "Angulo_Absoluto", acamMenuNEW, MenuName
        'fr.AddButton fr.CreateButtonBar("VBA Tab"), "Alinhar_Solido2.bmp", fr.LastMenuCommandID
        
    End With
    InitAlphacamAddIn = 0
    
End Function

Sub Angulo_Relativo() 'Medida considerando primeiro element como origem

Dim Drw As Drawing

Set Drw = App.ActiveDrawing

Drw.Font = "TArial"

Dim P1 As Path

Do

Set P1 = Drw.UserSelectOneGeo("ANGULOS: SELECIONE A GEOMETRIA")

If Not (P1 Is Nothing) Then

Dim h As Double

h = (P1.MaxXL - P1.MinXL + P1.MaxYL - P1.MinYL) / 80

Dim E As Element

For Each E In P1.Elements

Dim Ang As Double

Ang = E.AngleToElement(E.GetNext)

 'Chr(176) is the degrees symbol in a TrueType font

Dim S As String

S = Format(Ang, "0.00") & Chr(176)

 'Draw the text and loop for each path in the

 'returned collection, marking it as dimension

Dim P2 As Path

For Each P2 In Drw.CreateText(S, E.EndXL + h, E.EndYL + h, h)

P2.Dimension = True

P2.Redraw

Next P2

Next E

End If

Loop Until (P1 Is Nothing)

End Sub

Public Sub Angulo_Absoluto() ' Considerando o eixo x como referência
Dim Drw As Drawing
    Set Drw = App.ActiveDrawing
    Drw.Font = "TArial"

    Dim P1 As Path

    Do
         'Selecionar a linha
        Set P1 = Drw.UserSelectOneGeo("ANGULOS: SELECIONE UMA LINHA DE REFERÊNCIA")

        If Not (P1 Is Nothing) Then
            Dim h As Double
            h = (P1.MaxXL - P1.MinXL + P1.MaxYL - P1.MinYL) / 80

            Dim E As Element
            For Each E In P1.Elements
                Dim Ang As Double
                Dim DeltaX As Double, DeltaY As Double

                 'Calcular diferença das coordenadas em relação à origem (0,0)
                DeltaX = E.EndXL - E.StartXL
                DeltaY = E.EndYL - E.StartYL

                 'Calcular ângulo usando Atan2
                Ang = Atn(DeltaY / DeltaX) * 180 / 3.14159265358979

                 'Ajustar para os quatro quadrantes
                If DeltaX < 0 Then
                    Ang = Ang + 180
                ElseIf DeltaY < 0 Then
                    Ang = Ang + 360
                End If

                 'Formatar ângulo
                Dim S As String
                S = Format(Ang, "0.00") & Chr(176) ' O símbolo de graus (°)

                 'Criar a anotação de texto no desenho
                Dim P2 As Path
                For Each P2 In Drw.CreateText(S, E.EndXL + h, E.EndYL + h, h)
                    P2.Dimension = True
                    P2.Redraw
                Next P2
            Next E
        End If
    Loop Until (P1 Is Nothing)
End Sub
