VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConsolidateUI 
   Caption         =   "Entrada de matrices"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   OleObjectBlob   =   "ConsolidateUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConsolidateUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim is_local As Boolean
Private Sub AddElemMatrix()
    Dim ref_adress As String
    ref_adress = Mid(ReferenceInput.Value, InStr(1, ReferenceInput.Value, "!") + 1, 100)
    Dim i As Integer
    Dim should_add As Boolean
    should_add = True
    For i = 0 To LocalMatricesList.ListCount - 1
        If LocalMatricesList.List(i) = ref_adress Then should_add = False
    Next
    If should_add Then
        If ref_adress = "" Then
            MsgBox "Referencia vacía"
        Else
            LocalMatricesList.AddItem ref_adress
        End If
    Else
        MsgBox "Referencia repetida"
    End If
    ReferenceInput.Value = ""
End Sub

Private Sub CosolidateButton_Click()
    Dim elem_matrices() As Range
    Dim global_matrix As Range
    ReDim elem_matrices(1 To LocalMatricesList.ListCount)
    Dim i As Integer
    For i = 0 To LocalMatricesList.ListCount - 1
        Set elem_matrices(i + 1) = Range(LocalMatricesList.List(i))
    Next
    Set global_matrix = Range(GlobalMatrixRef.Caption)
    ConsolidateMatrix elem_matrices, global_matrix
End Sub

Private Sub DefineGlobalMatrix()
    Dim ref_adress As String
    ref_adress = Mid(ReferenceInputGlobal.Value, InStr(1, ReferenceInputGlobal.Value, "!") + 1, 100)
    GlobalMatrixRef.Caption = ref_adress
    ReferenceInputGlobal.Value = ""
End Sub

Private Sub DeleteItem_Click()
    If LocalMatricesList.ListIndex <> -1 Then
        LocalMatricesList.RemoveItem LocalMatricesList.ListIndex
    Else
        MsgBox "Ningún elemento seleccionado"
    End If
'    Stop
End Sub

Private Sub ReferenceInput_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 13 Then
        AddElemMatrix
    End If
End Sub

Private Sub ReferenceInputGlobal_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 13 Then
        DefineGlobalMatrix
    End If
End Sub

Private Sub UserForm_Click()

End Sub
