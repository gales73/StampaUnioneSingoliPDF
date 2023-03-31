Attribute VB_Name = "Utility"


Sub StampaUnioneSingoliPDF()
    Dim masterDoc As Document, singleDoc As Document, lastRecordNum As Long, nomeCampo1 As String, nomeCampo2 As String, inizFile As String, nomeFile As String
    
    inizFile = "Convocazione per "
    nomeCampo1 = "Cognome"
    nomeCampo2 = "Nome"
    Set masterDoc = ActiveDocument
    masterDoc.MailMerge.DataSource.ActiveRecord = wdLastRecord
    lastRecordNum = masterDoc.MailMerge.DataSource.ActiveRecord
    masterDoc.MailMerge.DataSource.ActiveRecord = wdFirstRecord
    'Seleziona la cartella di destinazione
      With Application.FileDialog(msoFileDialogFolderPicker)
          .Title = "Seleziona la cartella di destinazione"
          If .Show = -1 Then
              ExportFolder = .SelectedItems(1)
          Else
              Exit Sub
          End If
      End With
        
    Do While lastRecordNum > 0
        masterDoc.MailMerge.Destination = wdSendToNewDocument
        masterDoc.MailMerge.DataSource.FirstRecord = masterDoc.MailMerge.DataSource.ActiveRecord
        masterDoc.MailMerge.DataSource.LastRecord = masterDoc.MailMerge.DataSource.ActiveRecord
        masterDoc.MailMerge.Execute False
        Set singleDoc = ActiveDocument
        
        'Definisce il nome del file
        nomeFile = masterDoc.MailMerge.DataSource.DataFields(nomeCampo1).Value & " " & masterDoc.MailMerge.DataSource.DataFields(nomeCampo2).Value & " " & inizFile
		
        'singleDoc.SaveAs2 FileName:=ExportFolder & Application.PathSeparator  & nomeFile & ".docx", FileFormat:=wdFormatXMLDocument
                                                                   
        singleDoc.ExportAsFixedFormat OutputFileName:=ExportFolder & Application.PathSeparator & nomeFile & ".pdf", ExportFormat:=wdExportFormatPDF                                                                              ' Export "singleDoc" as a PDF with the details provided in the PdfFolderPath and PdfFileName fields in the MailMerge data
        
        singleDoc.Close False
        If masterDoc.MailMerge.DataSource.ActiveRecord >= lastRecordNum Then
            lastRecordNum = 0
        Else
            masterDoc.MailMerge.DataSource.ActiveRecord = wdNextRecord
        End If

    Loop
End Sub








