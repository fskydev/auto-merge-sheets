Attribute VB_Name = "merge"
Sub AutoMerge()
    Dim fnameList, fnameCurFile As Variant
    Dim wksCurSheet As Worksheet, wbkSrcSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
    Dim index As Integer, countSheets As Integer
    
    Dim companyTable As Variant
    Const columns = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI"
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        .ButtonName = "Select"
        If .Show = -1 Then
            sFolder = .SelectedItems(1)
            fnameList = GetFileNames(sFolder)
            
            If (vbBoolean <> VarType(fnameList) And Not IsEmpty(fnameList)) Then     'Not fnameList = Empty Or   And (fnameList.Count > 0)
                If (UBound(fnameList) > 0) Then
                    Dim i As Integer
                    Dim labels(34) As String
                    For i = 0 To 34
                        Dim j As Integer
                        j = i + 1
                        labels(i) = Utf8StringFromUtf8Bytes(j)
                    Next i
                    
                    letters = Split(columns, ",")
                    countSheets = 0
                    
                    Application.ScreenUpdating = False
                    Application.Calculation = xlCalculationManual
                    
                    Set wbkCurBook = ActiveWorkbook
                    Set wbkCurSheet = wbkCurBook.Sheets(1)
                    wbkCurSheet.Activate
                    wbkCurSheet.Range("A1:AI1").value = labels
                    
                    For Each fnameCurFile In fnameList
                        Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
                        For Each wbkSrcSheet In wbkSrcBook.Sheets
                              countSheets = countSheets + 1
                              Dim rngFound As Range, valueCell As Range
                          
                              With wbkSrcSheet.Cells
                                  Dim FilePath, FileOnly, PathOnly, unique As String
                                  FileOnly = wbkSrcBook.Name
                            
                                  unique = Split(FileOnly, ".")(0)
                                  Let IRange = letters(0) & (countSheets + 1)
                                  wbkCurSheet.Range(IRange).value = unique
                                  Let snrange = letters(1) & (countSheets + 1)
                                  wbkCurSheet.Range(snrange).value = wbkSrcSheet.Name
                                  
                                  For index = 3 To (UBound(labels) - LBound(labels) + 1)
                                      Let srange = letters(index - 1) & (countSheets + 1)
                                      Set rngFound = .Find(labels(index - 1))
                                      If Not rngFound Is Nothing Then
                                          If (rngFound.value <> "") Then
                                              Set valueCell = wbkSrcSheet.Cells(rngFound.Row, rngFound.Column + 1)
                                              If (StrComp(letters(index - 1), "F", vbTextCompare) = 0) Then
                                                  Debug.Print (valueCell.value)
                                                  wbkCurSheet.Range(srange).value = Format(valueCell.value, "yyyy " + Utf8YearLetter() + " mm " + Utf8MonthLetter())
                                              Else
                                                  wbkCurSheet.Range(srange).value = valueCell.value
                                              End If
                                          End If
                                      End If
                                  Next index
                              End With
                        Next
                        wbkSrcBook.Close SaveChanges:=False
                    Next
                    wbkCurSheet.columns("A:AI").AutoFit
                    'wbkCurSheet.Range("A1").EntireRow.Insert
                End If
            Else
                MsgBox "No files selected", Title:="Merge Excel files"
            End If
        End If
    End With
End Sub
