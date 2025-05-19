Attribute VB_Name = "Pleitos"
' ----- Version -----
'        1.2.0
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
        
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim tblRow As ListRow
    Dim newID As String
    Dim userResponse As VbMsgBoxResult
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    newID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("Esse aditivo já foi cadastrado. Deseja sobrescrever?", vbYesNoCancel + vbQuestion, "Confirmação")

        Select Case userResponse
            Case vbYes
                newID = Val(newID) ' Use ComboBoxID.Value as new ID
                ' Search for the ID in the first column of the table
                Set tblRow = dadosTable.ListRows(dadosTable.ListColumns(1).DataBodyRange.Find(What:=newID, LookAt:=xlWhole).Row - dadosTable.DataBodyRange.Row + 1)
            Case vbNo
                ' Proceed with generating new ID
                newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(1).DataBodyRange) + 1
                wsForm.OLEObjects("ComboBoxID").Object.Value = newID
                ' Add a new row to the table
                Set tblRow = dadosTable.ListRows.Add
            Case vbCancel
                Exit Sub ' Exit without saving
        End Select
    Else
        If dadosTable.ListColumns(colMap("ID")).DataBodyRange Is Nothing Then
            newID = 1
        Else
            newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(colMap("ID")).DataBodyRange) + 1
        End If
        
        wsForm.OLEObjects("ComboBoxID").Object.Value = newID
        
        wsForm.OLEObjects("ComboBoxName").Object.Value = wsForm.Range("B6").Value & " - " & wsForm.Range("B10").Value & " - " & wsForm.Range("D6").Value
        
        ' Add a new row to the table
        Set tblRow = dadosTable.ListRows.Add
    End If
    
    ' Assign values to the new row
    With tblRow.Range
        ' Set new ID
        .Cells(1, colMap("ID")).Value = newID ' First column value
        
        ' Read column B values
        .Cells(1, colMap("Obra")).Value = wsForm.Range("B6").Value
        .Cells(1, colMap("Cliente")).Value = wsForm.Range("B10").Value
        .Cells(1, colMap("Tipo")).Value = wsForm.Range("B14").Value
        .Cells(1, colMap("PM")).Value = wsForm.Range("B18").Value
        .Cells(1, colMap("PEP")).Value = wsForm.Range("B22").Value
        
        ' Read column D values
        .Cells(1, colMap("Descrição")).Value = wsForm.Range("D6").Value
        .Cells(1, colMap("Justificativa")).Value = wsForm.Range("D10").Value
        .Cells(1, colMap("Prestador")).Value = wsForm.Range("D14").Value
        .Cells(1, colMap("Valor")).Value = wsForm.Range("D18").Value
        
        ' Read column F values
        .Cells(1, colMap("Status")).Value = wsForm.Range("F6").Value
        .Cells(1, colMap("Data")).Value = "" 'Clear date if ovewriten in case an e-mail was already sent
        .Cells(1, colMap("Observações")).Value = wsForm.Range("F10").Value
        
    End With
    
    ' MsgBox "Dados salvos com sucesso!", vbInformation
End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchName As String
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxName").Object.Value <> "" Then
        searchName = wsForm.OLEObjects("ComboBoxName").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the matching row
    Set foundRow = Nothing
    For Each cell In dadosTable.ListColumns(colMap("ID")).DataBodyRange
        If cell.Value & " - " & cell.Cells(1, colMap("Cliente")).Value & " - " & cell.Cells(1, colMap("Descrição")).Value = searchName Then
            Set foundRow = cell
            Exit For
        End If
    Next cell
    
    ' If Name is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "Nenhuma obra encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxID").Object.Value = foundRow.Value
    
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Obra")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PM")).Value
        .Range("B22").Value = foundRow.Cells(1, colMap("PEP")).Value
    
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Descrição")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Justificativa")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Prestador")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Valor")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("Observações")).Value
    End With
End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchID As Double
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxID").Object.Value <> "" Then
        searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(colMap("ID")).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Cells(1, colMap("ID")).Value & " - " & foundRow.Cells(1, colMap("Obra")).Value & " - " & foundRow.Cells(1, colMap("Descrição")).Value
        
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("ID")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Obra")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PM")).Value
        .Range("B22").Value = foundRow.Cells(1, colMap("PEP")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Descrição")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Justificativa")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Prestador")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Valor")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("Observações")).Value
    End With
End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)
    
    Dim wsForm As Worksheet
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    
    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados não foram salvos. Deseja limpá-los mesmo assim?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        .OLEObjects("ComboBoxID").Object.Value = ""
        .OLEObjects("ComboBoxName").Object.Value = ""
        .OLEObjects("ComboBoxName").Width = 123
        
        ' Read column B values
        .Range("B6").Value = ""
        .Range("B10").Value = ""
        .Range("B14").Value = ""
        .Range("B18").Value = ""
        .Range("B22").Value = ""
        
        ' Read column D values
        .Range("D6").Value = ""
        .Range("D10").Value = ""
        .Range("D14").Value = ""
        .Range("D18").Value = ""
    
        ' Read column F values
        .Range("F6").Value = ""
        .Range("F10").Value = ""
    End With
End Sub

Public Function GetColumnHeadersMapping() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "ID", 1
    headers.Add "Obra", 2
    headers.Add "Cliente", 3
    headers.Add "Tipo", 4
    headers.Add "PM", 5
    headers.Add "PEP", 6
    headers.Add "Descrição", 7
    headers.Add "Justificativa", 8
    headers.Add "Prestador", 9
    headers.Add "Valor", 10
    headers.Add "Status", 11
    headers.Add "Data", 12
    headers.Add "Observações", 13

    Set GetColumnHeadersMapping = headers
End Function
