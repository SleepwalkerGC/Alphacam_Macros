Function InitAlphacamAddIn(acamversion As Long) As Integer
    
    Dim frm As Frame
    Set frm = App.Frame
    
    frm.AddMenuItem2 "CreateRectangleitem", "CreateRecButton", acamMenuNEW, "VBA Tab" '( frm.AddMenuItem2 name of button in AC interface, Name of function, New Menu Bar, Name of Tab )
    frm.AddButton frm.CreateButtonBar("VBA Tab"), "CreateRectangle.bmp", frm.LastMenuCommandID
    
    InitAlphacamAddIn = 0
    
End Function
 
Function CreateRecButton()
    
    Dim myUserForm As frmMain ' name of form
    
    Set myUserForm = New frmMain 'name of form
    myUserForm.Show
         
End Function

'FORM ===============================================
Private Sub cmdCreate_Click()

    Dim Drw As Drawing
    Dim P As Path
    Dim Lyr As Layer
   
    
    
        Set Drw = App.ActiveDrawing
        Set Lyr = Drw.CreateLayer("Bordas")
        Lyr.Color = acamRED
        
            
        Drw.SetLayer Lyr
         
            Set P = Drw.CreateRectangle(0, 0, txtComprimento.Value, txtLargura.Value)
          
            
            
    
         Set Lyr = Drw.CreateLayer("Chapa_Útil")
         Lyr.Color = acamCYAN
         Lyr.LineType = acamLineDOT
         
         Drw.SetLayer Lyr
        
            Set P = Drw.CreateRectangle(txtRail.Value, txtRail.Value, txtComprimento.Value - txtRail.Value, txtLargura.Value - txtRail.Value)
            'Const acamViewTOP = 14 ' vista Top
            'Drw.ZoomAll
            'Drw.Refresh
            
            'Move view to Top after execute
            Dim Vista As ViewWindow ' set var to Top View
            Set Vista = App.ActiveDrawing.CurrentViewWindow
            Vista.View = acamViewTOP
            
            
                    
         Drw.SetLayer Nothing
         
         
         Unload Me ' Close Macro
         

End Sub
'End FORM ===============================================================================