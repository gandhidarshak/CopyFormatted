Attribute VB_Name = "Module1"

'-------------------------------------------------------------------------------
' Summary: VBA macros to copy data from excel into ASCII files in formats like
' table, csv, yaml, etc. More: https://github.com/gandhidarshak/CopyFormatted
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Externally visible sub-routines.  Don't change these sub-routines if possible
' as people may have made shortcuts around these
'-------------------------------------------------------------------------------
Sub CopyAsTable()
   InternalCopyAsTable
End Sub
Sub CopyAsCSV()
   InternalCopyAsCSV
End Sub
Sub CopyAsYaml()
   InternalCopyAsYaml
End Sub

'-------------------------------------------------------------------------------
' Internal Function for Copy to Table
'-------------------------------------------------------------------------------
Function InternalCopyAsTable()
   ' Load arrayData from excle
   Dim arrayData() As String
   Dim columnSize() As Integer
   Dim rCnt As Integer
   Dim cCnt As Integer
   Call createArray(arrayData, columnSize, rCnt, cCnt)

   ' Create Clipboard String by using arrayData and formatting them using columnSize
   Dim clipStr As String
   clipStr = ""
   clipStr = clipStr + createSeparaterRow(columnSize, rCnt, cCnt)
   Dim firstRowHeader As Boolean
   firstRowHeader = True

   For rIter = 0 To rCnt - 1
      Dim rowUsed As Boolean
      rowUsed = False
      For cIter = 0 To cCnt - 1
         If columnSize(rIter, cIter) <> -1 Then
            celVal = arrayData(rIter, cIter)
            celVal = AlignAndPad(celVal, columnSize(rIter, cIter))
            clipStr = clipStr + "| " + celVal + " "
            rowUsed = True
         End If
      Next
      If rowUsed = True Then
         clipStr = clipStr + "|" + vbNewLine
         If firstRowHeader = True Then
            clipStr = clipStr + createSeparaterRow(columnSize, rCnt, cCnt)
            firstRowHeader = False
         End If
      End If
   Next
   clipStr = clipStr + createSeparaterRow(columnSize, rCnt, cCnt)

   ' Put the clipStr into clipboard
   Call pasteToClipboard(clipStr)
End Function

'-------------------------------------------------------------------------------
' Internal Function for Copy to CSV
'-------------------------------------------------------------------------------
Function InternalCopyAsCSV()
   ' Load arrayData from excle
   Dim arrayData() As String
   Dim columnSize() As Integer
   Dim rCnt As Integer
   Dim cCnt As Integer
   Call createArray(arrayData, columnSize, rCnt, cCnt)

   ' Create Clipboard String by using arrayData and formatting them using columnSize
   Dim clipStr As String
   clipStr = ""

   For rIter = 0 To rCnt - 1
      Dim rowUsed As Boolean
      rowUsed = False
      For cIter = 0 To cCnt - 1
         If columnSize(rIter, cIter) <> -1 Then
            celVal = arrayData(rIter, cIter)
            celVal = AlignAndPad(celVal, columnSize(rIter, cIter))
            If cIter = 0 Then
               clipStr = clipStr + celVal + " "
            Else
               clipStr = clipStr + ", " + celVal + " "
            End If
            rowUsed = True
         End If
      Next
      If rowUsed = True Then
         clipStr = clipStr + vbNewLine
      End If
   Next

   ' Put the clipStr into clipboard
   Call pasteToClipboard(clipStr)
End Function

'-------------------------------------------------------------------------------
' Internal Function for Copy to Yaml
'-------------------------------------------------------------------------------
Function InternalCopyAsYaml()
   ' Load arrayData from excle
   Dim arrayData() As String
   Dim columnSize() As Integer
   Dim rCnt As Integer
   Dim cCnt As Integer
   Call createArray(arrayData, columnSize, rCnt, cCnt)

   ' Create Clipboard String by using arrayData and formatting them using columnSize
   Dim clipStr As String
   clipStr = ""

   For rIter = 0 To rCnt - 1
      Dim rowUsed As Boolean
      rowUsed = False
      For cIter = 0 To cCnt - 1
         If columnSize(rIter, cIter) <> -1 Then
            celVal = arrayData(rIter, cIter)
            If cIter = 0 Then
               celVal = Replace(celVal, " ", "_")
            End If
            celVal = AlignAndPad(celVal, columnSize(rIter, cIter))
            If cIter = 0 Then
               clipStr = clipStr + celVal + " "
            ElseIf cIter = 1 Then
               clipStr = clipStr + ": [ " + celVal + " "
            Else
               clipStr = clipStr + ", " + celVal + " "
            End If
            rowUsed = True
         End If
      Next
      If rowUsed = True Then
         clipStr = clipStr + "]" + vbNewLine
      End If
   Next

   ' Put the clipStr into clipboard
   Call pasteToClipboard(clipStr)
End Function

'-------------------------------------------------------------------------------
' Paste to Clipboard API
' If you see any Library Errors, open your VBA editor. Click Tools > References.
' Check the box next to “Microsoft Forms 2.0 Object Library.”
'-------------------------------------------------------------------------------
Function pasteToClipboard(ByRef clipStr As String)
   Dim clipboard As MSForms.DataObject
   Set clipboard = New MSForms.DataObject
   clipboard.SetText clipStr
   clipboard.PutInClipboard
End Function

'-------------------------------------------------------------------------------
' Row separator for 1st, 2nd and Last rows of table.
'-------------------------------------------------------------------------------
Function createSeparaterRow(ByRef columnSize() As Integer, _
   ByRef rCnt As Integer, _
   ByRef cCnt As Integer) As String
   For rIter = 0 To rCnt - 1
      Dim rowUsed As Boolean
      rowUsed = False
      For cIter = 0 To cCnt - 1
         If columnSize(rIter, cIter) <> -1 Then
            Dim celVal As String
            celVal = String(columnSize(rIter, cIter), "-")
            createSeparaterRow = createSeparaterRow + "+-" + celVal + "-"
            rowUsed = True
         End If
      Next
      If rowUsed = True Then
         createSeparaterRow = createSeparaterRow + "+" + vbNewLine
         Exit For
      End If
   Next
End Function

'-------------------------------------------------------------------------------
' Convert selected cells into an array, retur column size for each cell
'-------------------------------------------------------------------------------
Function createArray(ByRef arrayData() As String, _
                     ByRef columnSize() As Integer, _
                     ByRef rCnt As Integer, _
                     ByRef cCnt As Integer)
   
   ' Get user selected range
   Dim selRange As Range
   Set selRange = Application.Selection

   '----------------------------------------------------------------------------
   ' Depending on order of selections selRange may be weird.  So to find out
   ' min/max columns/rows, safest way is to iterate over all Cells.
   '----------------------------------------------------------------------------
   Dim rMin, rMax, cMin, cMax As Integer
   rMin = 1000000
   cMin = 1000000
   rMax = -1
   cMax = -1

   On Error Resume Next
   For Each selCel In selRange.Cells
      If rMin > selCel.Row Then
         rMin = selCel.Row
      End If
      If cMin > selCel.Column Then
         cMin = selCel.Column
      End If
      If rMax < selCel.Row Then
         rMax = selCel.Row
      End If
      If cMax < selCel.Column Then
         cMax = selCel.Column
      End If
   Next selCel
   rCnt = rMax - rMin + 1
   cCnt = cMax - cMin + 1

   ' Store the selected data in array with column size in-terms of char count
   ReDim arrayData(rCnt, cCnt)
   ReDim columnSize(rCnt, cCnt)
   Dim celVal As String

   On Error Resume Next
   For Each selCel In selRange.Cells
      celVal = LTrim(RTrim(selCel.Text))
      arrayData(selCel.Row - rMin, selCel.Column - cMin) = celVal
      columnSize(selCel.Row - rMin, selCel.Column - cMin) = Len(celVal)
   Next selCel

   Dim rIter, cIter As Integer
   Dim maxSize As Integer
   ' Detect empty row first as that will help us set column width when we detect
   ' empty columns
   For rIter = 0 To rCnt - 1
      maxSize = 0
      For cIter = 0 To cCnt - 1
         If maxSize < columnSize(rIter, cIter) Then
            maxSize = columnSize(rIter, cIter)
         End If
      Next
      If maxSize = 0 Then
         For cIter = 0 To cCnt - 1
            columnSize(rIter, cIter) = -1
         Next
      End If
   Next
   For cIter = 0 To cCnt - 1
      maxSize = 0
      For rIter = 0 To rCnt - 1
         If maxSize < columnSize(rIter, cIter) Then
            maxSize = columnSize(rIter, cIter)
         End If
      Next
      If maxSize = 0 Then
         For rIter = 0 To rCnt - 1
            columnSize(rIter, cIter) = -1
         Next
      Else
         For rIter = 0 To rCnt - 1
            columnSize(rIter, cIter) = maxSize
         Next
      End If
   Next
End Function

'-------------------------------------------------------------------------------
' Alignment and padding functions
'-------------------------------------------------------------------------------
Function LeftAlignedPad(orgStr As String, size As Integer) As String
   size = WorksheetFunction.Max(size, Len(orgStr))
   LeftAlignedPad = Left(orgStr + String(size, " "), size)
End Function

Function RightAlignedPad(orgStr As String, size As Integer) As String
   size = WorksheetFunction.Max(size, Len(orgStr))
   RightAlignedPad = Right(String(size, " ") + orgStr, size)
End Function

Function CenterAlignedPad(orgStr As String, size As Integer) As String
   CenterAlignedPad = LeftAlignedPad(orgStr, Int(Len(orgStr) + size) / 2)
   CenterAlignedPad = RightAlignedPad(CenterAlignedPad, size)
End Function

Function AlignAndPad(ByVal orgStr As String, size As Integer) As String
   If IsDate(orgStr) Then
      AlignAndPad = RightAlignedPad(orgStr, size)
   ElseIf IsNumeric(orgStr) Then
      AlignAndPad = RightAlignedPad(orgStr, size)
   Else
      AlignAndPad = LeftAlignedPad(orgStr, size)
   End If
End Function

