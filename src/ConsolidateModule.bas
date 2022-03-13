Attribute VB_Name = "ConsolidateModule"
Option Explicit

Public Sub ConsolidateMatrix(elem_matrices() As Range, global_matrix As Range)

Dim i As Integer, j As Integer

'Dim elem_matrices(1 To 5) As Range
'
'Set elem_matrices(1) = Range("E112:K118")
'Set elem_matrices(2) = Range("E128:K134")
'Set elem_matrices(3) = Range("E144:K150")
'Set elem_matrices(4) = Range("E160:K166")
'Set elem_matrices(5) = Range("E176:K182")
'
'Dim global_matrix As Range
'
'Set global_matrix = Range("H190:Z208")


' Clear the global matrix
For i = 2 To global_matrix.Rows.count
    For j = 2 To global_matrix.Columns.count
        global_matrix.Cells(i, j).Value = ""
    Next
Next

' Iterate over the element matrices
Dim elem_matrix As Variant
For Each elem_matrix In elem_matrices
    ' Iterate over each position of the matrix
    For i = 2 To elem_matrix.Rows.count
        For j = 2 To elem_matrix.Columns.count
            Dim current_cell As Range
            Set current_cell = elem_matrix.Cells(i, j)
            Dim row_to_fill As Integer, col_to_fill As Integer
            row_to_fill = CInt(elem_matrix.Cells(i, 1).Value) + 1
            col_to_fill = CInt(elem_matrix.Cells(1, j).Value) + 1
            
            If col_to_fill <= global_matrix.Columns.count And row_to_fill <= global_matrix.Rows.count Then
            
                If global_matrix.Cells(row_to_fill, col_to_fill).Value = "" Then
                    global_matrix.Cells(row_to_fill, col_to_fill).Formula = "=" & current_cell.Address
                Else
                    global_matrix.Cells(row_to_fill, col_to_fill).Formula = global_matrix.Cells(row_to_fill, col_to_fill).Formula & "+" & current_cell.Address
                End If
            
            End If
            
        Next
    Next
Next

' Fill with zeros the empty ones
For i = 2 To global_matrix.Rows.count
    For j = 2 To global_matrix.Columns.count
        If global_matrix.Cells(i, j).Value = "" Then
            global_matrix.Cells(i, j).Value = 0
        End If
    Next
Next

End Sub

Public Sub RunUI()
Attribute RunUI.VB_ProcData.VB_Invoke_Func = "Z\n14"
    ConsolidateUI.Show
End Sub

