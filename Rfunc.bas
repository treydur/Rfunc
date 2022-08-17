Attribute VB_Name = "Rfunc"
Sub fillTable()

'Trey Durden
'tdurden8@gatech.edu
'https://github.com/treydur

'For this program to work named ranges must be used
'This allows the program to be dynamic, so if rows are inserted it will still work
'DataRange: the cells in table from subject to credit hours excluding header
'RecoRange1: the cells from course to credit hours column in left table excluding header
'RecoRange2: the cells from course to credit hours column in right table excluding header

Dim DataRange As Variant
Dim RecoRange1 As Variant
Dim RecoRange2 As Variant
Dim Grade As String
Dim Irow As Long
Dim Icol As Integer
Dim numRows As Long
Dim i As Integer
Dim k As Integer
Dim n As Integer

Range("RecoRange1").ClearContents
Range("RecoRange2").ClearContents

'Read all the values at once into arrays
DataRange = Range("DataRange").Value2
RecoRange1 = Range("RecoRange1").Value2
RecoRange2 = Range("RecoRange2").Value2

i = 1
n = UBound(RecoRange1)
k = n + UBound(RecoRange2)

For Irow = 1 To UBound(DataRange)
    Grade = DataRange(Irow, 3)
    If Grade = "R" Or Grade = "r" And i <= k Then
        If i <= n Then
            RecoRange1(i, 1) = DataRange(Irow, 1) 'subject
            RecoRange1(i, 2) = DataRange(Irow, 2) 'course number
            RecoRange1(i, 3) = DataRange(Irow, 4) 'credit hours
        End If
        
        'Left table is full so use right table
        If i > n Then
            t = i - n
            RecoRange2(t, 1) = DataRange(Irow, 1)
            RecoRange2(t, 2) = DataRange(Irow, 2)
            RecoRange2(t, 3) = DataRange(Irow, 4)
        End If
        i = i + 1
    End If
Next Irow

'Write all the values from the arrays at once
Range("RecoRange1").Value2 = RecoRange1
Range("RecoRange2").Value2 = RecoRange2

End Sub
