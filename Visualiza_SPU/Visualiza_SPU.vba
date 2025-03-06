Public Function InitAlphacamAddIn(acamversion As Long) As Integer
    Dim fr As Frame
    Set fr = App.Frame
    With fr
        ' set up itemname  and menuname as new string variables
        Dim ItemName As String, MenuName As String
        ItemName = "SPU": MenuName = "VBA Tab"
        ' create the new menu
        .AddMenuItem2 ItemName, "AbrirArquivoARD", acamMenuNEW, MenuName
        fr.AddButton fr.CreateButtonBar("VBA Tab"), "Visualiza_SPU.bmp", fr.LastMenuCommandID
    
    End With
    InitAlphacamAddIn = 0

End Function

Public Sub AbrirArquivoARD()
    Dim Drw As Object
    Dim diretorioEspecifico As String
    Dim nomeArquivo As String
    Dim caminhoArquivo As String
    
    ' Diretório específico onde os arquivos .ard estão
    diretorioEspecifico = "\\server\index\Projetos\Programas\AlphaCAM\"
    
    ' Solicita o nome do arquivo ao usuário (sem a extensão)
    nomeArquivo = InputBox("Digite o nome do arquivo (sem extensão):")
    
    ' Verifica se o nome do arquivo foi inserido
    If nomeArquivo = "" Then
        MsgBox "Nome do arquivo não inserido.", vbExclamation
        Exit Sub
    End If
    
    ' Caminho completo do arquivo
    caminhoArquivo = diretorioEspecifico & nomeArquivo & ".ard"
    
    ' Verifica se o arquivo existe
    If Dir(caminhoArquivo) <> "" Then
        ' Tenta conectar ao Alphacam e abrir o arquivo
        On Error Resume Next
        Set Drw = App.Application
        On Error GoTo ErroAbertura
        
        ' Abre o arquivo no Alphacam
        Drw.OpenDrawing caminhoArquivo
        'Move view to Top after execute
            Dim Vista As ViewWindow ' set var to ISO View
            Set Vista = App.ActiveDrawing.CurrentViewWindow
            Vista.View = acamViewISO
        'MsgBox "Arquivo aberto com sucesso!", vbInformation
        Exit Sub
    Else
        ' Arquivo não encontrado
        MsgBox "Arquivo não encontrado no diretório especificado.", vbExclamation
    End If
    Exit Sub
    
ErroAbertura:
    MsgBox "Erro ao abrir o arquivo. Verifique o caminho e tente novamente.", vbCritical
    
    
End Sub