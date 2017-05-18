Attribute VB_Name = "Module1"
Option Explicit

Public Const X_length = 52
Public Const Y_length = 285
Public Const Row_cell_begin = 2
Public Const Column_cell_begin = 2
Public Const Row_cell_end = 286
Public Const Column_cell_end = 53

Dim CP1_Map(Column_cell_begin To Column_cell_end, Row_cell_begin To Row_cell_end) As String
Dim CP2_Map(Column_cell_begin To Column_cell_end, Row_cell_begin To Row_cell_end) As String
Dim Result_Map(Column_cell_begin To Column_cell_end, Row_cell_begin To Row_cell_end) As String

Sub TwoWaferMapCombine()

    Dim row, column As Long
    Dim Bin1_cnt As Long
    Dim temp_string As String
    
    Bin1_cnt = 0
    
    For column = Column_cell_begin To Column_cell_end
        For row = Row_cell_begin To Row_cell_end
            
            If IsEmpty(Worksheets("CP1").Cells(row, column).Value) Then
                CP1_Map(column, row) = " "
                CP2_Map(column, row) = " "
            Else
                CP1_Map(column, row) = CStr(Worksheets("CP1").Cells(row, column).Value)
                CP2_Map(column, row) = CStr(Worksheets("CP2").Cells(row, column).Value)
            
            End If
        Next row
    Next column
    
    For column = Column_cell_begin To Column_cell_end
        For row = Row_cell_begin To Row_cell_end
            
            If CP1_Map(column, row) = " " Then
                Result_Map(column, row) = "."
            Else
                If CP1_Map(column, row) = "1" And CP2_Map(column, row) = "1" Then
                    Result_Map(column, row) = "A"
                    Bin1_cnt = Bin1_cnt + 1
                Else
                    Result_Map(column, row) = "X"
                End If
            End If
            
        Next row
    Next column
    
    For row = Row_cell_begin To Row_cell_end
        temp_string = ""
        For column = Column_cell_begin To Column_cell_end
            temp_string = temp_string + Result_Map(column, row)
        Next column
        Worksheets("Test").Cells(row, 1).Value = temp_string
    Next row
    
    MsgBox "Pass quantity = " + CStr(Bin1_cnt)
    
End Sub
