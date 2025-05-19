VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBoxID_DropButtonClick()
    
    Dim colMap As Object
    Set colMap = Pleitos.GetColumnHeadersMapping()
    
    Dim wsDados As Worksheet
    Dim wsFormulário As Worksheet
    Dim dadosTable As ListObject
    Dim sortedKeys As Object
    Dim currentList As Object
    Dim cell As Range
    Dim i As Integer
    Dim isListSame As Boolean
    
    ' Define worksheets
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsForm = ThisWorkbook.Sheets("Formulário")

    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    If dadosTable Is Nothing Then Exit Sub ' Exit if table not found
    
    ' Create an ArrayList to store and sort unique values
    Set sortedKeys = CreateObject("System.Collections.ArrayList")
    
    ' Loop through column 1 values and add unique ones
    For Each cell In dadosTable.ListColumns(colMap("ID")).DataBodyRange
        If Not IsEmpty(cell.Value) And Not sortedKeys.Contains(cell.Value) Then
            sortedKeys.Add "" & cell.Value
        End If
    Next cell
    
    ' Sort the ArrayList alphabetically
    sortedKeys.Sort
    
    ' Ensure the first key is an empty string
    'If sortedKeys.Count > 0 Then
    '    sortedKeys.Insert 0, "" ' Insert empty string at the beginning
    'End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Retrieve the current list of ComboBox items
    Set currentList = CreateObject("System.Collections.ArrayList")
    With wsForm.OLEObjects("ComboBoxID").Object ' Change "ComboBoxID" to the correct name of your ComboBox
        For i = 0 To .ListCount - 1
            currentList.Add .List(i)
        Next i
    End With
    
    ' Compare the current ComboBox list with the new sorted list
    isListSame = (currentList.Count = sortedKeys.Count)
    If isListSame Then
        For i = 0 To sortedKeys.Count - 1
            If sortedKeys(i) <> currentList(i) Then
                isListSame = False
                Exit For
            End If
        Next i
    End If

    ' Exit if the list is already up-to-date
    If isListSame Then Exit Sub
    
    ' Populate Form Control ComboBox
    With wsForm.OLEObjects("ComboBoxID").Object
        .Clear ' Clear existing items
        For Each key In sortedKeys
            .AddItem key
        Next key
    End With
End Sub

Private Sub ComboBoxID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Check if the Enter key is pressed
    If KeyCode = vbKeyReturn Then
        RetrieveDataFromID
    End If
End Sub

Private Sub ComboBoxID_LostFocus()
    RetrieveDataFromID
End Sub

Private Sub ComboBoxName_DropButtonClick()

    Dim colMap As Object
    Set colMap = Pleitos.GetColumnHeadersMapping()
    
    Dim wsDados As Worksheet
    Dim wsFormulário As Worksheet
    Dim dadosTable As ListObject
    Dim sortedKeys As Object
    Dim currentList As Object
    Dim cell As Range
    Dim i As Integer
    Dim isListSame As Boolean
    Dim maxChar As Long
    
    ' Define worksheets
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsForm = ThisWorkbook.Sheets("Formulário")

    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    If dadosTable Is Nothing Then Exit Sub ' Exit if table not found
    
    ' Create an ArrayList to store and sort unique values
    Set sortedKeys = CreateObject("System.Collections.ArrayList")

    maxChar = 125 / 5
    
    ' Loop through column 1 values and add unique ones
    For Each cell In dadosTable.ListColumns(colMap("ID")).DataBodyRange
        If Not IsEmpty(cell.Value) And Not sortedKeys.Contains(cell.Value & " - " & cell.Cells(1, colMap("Obra")) & " - " & cell.Cells(1, colMap("Descrição"))) Then
            If maxChar < Len(cell.Value & " - " & cell.Cells(1, colMap("Obra")) & " - " & cell.Cells(1, colMap("Descrição"))) Then
                maxChar = Len(cell.Value & " - " & cell.Cells(1, colMap("Obra")) & " - " & cell.Cells(1, colMap("Descrição")))
            End If
            
            sortedKeys.Add cell.Value & " - " & cell.Cells(1, colMap("Obra")) & " - " & cell.Cells(1, colMap("Descrição"))
        End If
    Next cell
    
    ' Sort the ArrayList alphabetically
    sortedKeys.Sort
    
    ' Ensure the first key is an empty string
    'If sortedKeys.Count > 0 Then
    '    sortedKeys.Insert 0, "" ' Insert empty string at the beginning
    'End If
    
    wsForm.OLEObjects("ComboBoxName").Width = maxChar * 5
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Retrieve the current list of ComboBox items
    Set currentList = CreateObject("System.Collections.ArrayList")
    With wsForm.OLEObjects("ComboBoxName").Object ' Change "ComboBoxID" to the correct name of your ComboBox
        For i = 0 To .ListCount - 1
            currentList.Add .List(i)
        Next i
    End With
    
    ' Compare the current ComboBox list with the new sorted list
    isListSame = (currentList.Count = sortedKeys.Count)
    If isListSame Then
        For i = 0 To sortedKeys.Count - 1
            If sortedKeys(i) <> currentList(i) Then
                isListSame = False
                Exit For
            End If
        Next i
    End If
    
    ' Exit if the list is already up-to-date
    If isListSame Then Exit Sub
    
    ' Populate Form Control ComboBox
    With wsForm.OLEObjects("ComboBoxName").Object
        .Clear ' Clear existing items
        For Each key In sortedKeys
            .AddItem key
        Next key
    End With
    
End Sub

Private Sub ComboBoxName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Check if the Enter key is pressed
    If KeyCode = vbKeyReturn Then
        RetrieveDataFromName
    End If
End Sub

Private Sub ComboBoxName_LostFocus()
    RetrieveDataFromName
End Sub
