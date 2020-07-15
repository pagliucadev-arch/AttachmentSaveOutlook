VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub ProcessarAnexo(Email As MailItem)
    Dim DiretorioAnexos As String
    
    DiretorioAnexos = "C:\anexos"
    
    Dim MailID As String
    Dim Mail As Outlook.MailItem
    
    MailID = Email.EntryID
    Set Mail = Application.Session.GetItemFromID(MailID)
        
    For Each Anexo In Mail.Attachments
        If Right(Anexo.FileName, 3) = "txt" Then
            Anexo.SaveAsFile DiretorioAnexos & "\" & Anexo.FileName
            
            ElseIf Right(Anexo.FileName, 3) = "TXT" Then
            Anexo.SaveAsFile DiretorioAnexos & "\" & Anexo.FileName
        
        End If
    Next
    
    Set Mail = Nothing
End Sub

