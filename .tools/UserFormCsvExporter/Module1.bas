Option Explicit

Public ExportMode As Boolean

Private gStream As Object
Private CsvPath As String

Public Sub ExportUserFormCsvFiles()

    Dim candidateWb As Workbook
    Dim targetWb As Workbook
    Dim workbookChoices() As Workbook
    Dim workbookCount As Long
    Dim choice As Variant
    Dim prompt As String
    Dim formCount As Long

    workbookCount = 0

    For Each candidateWb In Application.Workbooks
        If Not candidateWb.IsAddin And candidateWb.Name <> ThisWorkbook.Name Then
            workbookCount = workbookCount + 1
        End If
    Next candidateWb

    If workbookCount = 0 Then
        MsgBox "対象ブックが見つかりません。対象のExcelブックを開いてから実行してください。", vbInformation
        Exit Sub
    End If

    ReDim workbookChoices(1 To workbookCount)
    prompt = "UserForm情報を取得するブックの番号を入力してください。" & vbCrLf & vbCrLf
    workbookCount = 0

    For Each candidateWb In Application.Workbooks
        If Not candidateWb.IsAddin And candidateWb.Name <> ThisWorkbook.Name Then
            workbookCount = workbookCount + 1
            Set workbookChoices(workbookCount) = candidateWb
            prompt = prompt & workbookCount & ". " & candidateWb.Name & vbCrLf
        End If
    Next candidateWb

    choice = Application.InputBox(prompt, "対象ブックの選択", Type:=1)

    If VarType(choice) = vbBoolean Then Exit Sub
    If CDbl(choice) <> Fix(CDbl(choice)) Then
        MsgBox "一覧にある整数の番号を入力してください。", vbExclamation
        Exit Sub
    End If
    If CDbl(choice) < 1 Or CDbl(choice) > workbookCount Then
        MsgBox "一覧にある番号を入力してください。", vbExclamation
        Exit Sub
    End If

    Set targetWb = workbookChoices(CLng(choice))

    If Len(targetWb.Path) = 0 Then
        MsgBox "対象ブックがまだ保存されていません。先に保存してから実行してください。", vbExclamation
        Exit Sub
    End If

    On Error GoTo ExportError
    formCount = ExportWorkbookForms(targetWb)

    If formCount = 0 Then
        MsgBox "対象ブックにUserFormはありません。", vbInformation
    Else
        MsgBox "出力完了: " & formCount & "件" & vbCrLf & targetWb.Path, vbInformation
    End If

    Exit Sub

ExportError:
    CloseCurrentStream
    MsgBox "UserForm情報の出力に失敗しました。" & vbCrLf & _
           "Excelの「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定と、対象プロジェクトの保護状態を確認してください。" & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbExclamation

End Sub

Public Function ExportUserFormCsvFilesForPath(ByVal TargetWorkbookPath As String) As Long

    Dim targetWb As Workbook
    Dim openedByExporter As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo ExportPathError

    If Len(Trim$(TargetWorkbookPath)) = 0 Then
        Err.Raise vbObjectError + 513, "ExportUserFormCsvFilesForPath", "対象ブックのパスが空です。"
    End If

    Set targetWb = FindOpenWorkbook(TargetWorkbookPath)
    If targetWb Is Nothing Then
        Set targetWb = Application.Workbooks.Open( _
            Filename:=TargetWorkbookPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            AddToMru:=False)
        openedByExporter = True
    End If

    If targetWb Is ThisWorkbook Or targetWb.IsAddin Then
        Err.Raise vbObjectError + 514, "ExportUserFormCsvFilesForPath", "アドイン自身は対象ブックにできません。"
    End If
    If Len(targetWb.Path) = 0 Then
        Err.Raise vbObjectError + 515, "ExportUserFormCsvFilesForPath", "対象ブックが保存されていません。"
    End If

    ExportUserFormCsvFilesForPath = ExportWorkbookForms(targetWb)

    If openedByExporter Then targetWb.Close SaveChanges:=False
    Exit Function

ExportPathError:
    errorNumber = Err.Number
    errorText = Err.Description
    CloseCurrentStream

    On Error Resume Next
    If openedByExporter Then targetWb.Close SaveChanges:=False
    On Error GoTo 0

    Err.Raise errorNumber, "ExportUserFormCsvFilesForPath", errorText

End Function

Private Function FindOpenWorkbook(ByVal TargetWorkbookPath As String) As Workbook

    Dim candidateWb As Workbook

    For Each candidateWb In Application.Workbooks
        If StrComp(candidateWb.FullName, TargetWorkbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = candidateWb
            Exit Function
        End If
    Next candidateWb

End Function

Private Function ExportWorkbookForms(ByVal targetWb As Workbook) As Long

    Dim vbComp As Object
    Dim formCount As Long
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo ExportError

    For Each vbComp In targetWb.VBProject.VBComponents
        If vbComp.Type = 3 Then
            formCount = formCount + 1
            CsvPath = targetWb.Path & Application.PathSeparator & vbComp.Name & ".csv"
            CreateCsvFile CsvPath
            WriteHeader
            ExportControlsRecursive vbComp.Designer.Controls, vbComp.Name, ""
            SaveCsv
            Set gStream = Nothing
        End If
    Next vbComp

    ExportWorkbookForms = formCount
    Exit Function

ExportError:
    errorNumber = Err.Number
    errorText = Err.Description
    CloseCurrentStream
    Err.Raise errorNumber, "ExportWorkbookForms", errorText

End Function

Private Sub CloseCurrentStream()

    On Error Resume Next
    If Not gStream Is Nothing Then
        If gStream.State = 1 Then gStream.Close
    End If
    Set gStream = Nothing
    On Error GoTo 0

End Sub

Private Sub CreateCsvFile(ByVal FilePath As String)

    Set gStream = CreateObject("ADODB.Stream")

    With gStream
        .Charset = "UTF-8"
        .Open
    End With

End Sub

Private Sub SaveCsv()

    gStream.SaveToFile CsvPath, 2
    gStream.Close

End Sub

Private Sub WriteHeader()

    gStream.WriteText _
        "FormName,ParentControl,ControlName,ControlType,Caption,Top,Left,Width,Height,TabIndex" & _
        vbCrLf

End Sub

Private Sub ExportControlsRecursive( _
        ByVal Controls As Object, _
        ByVal FormName As String, _
        ByVal ParentName As String)

    Dim ControlArray() As Object
    Dim ctl As Object
    Dim i As Long
    Dim controlCount As Long

    controlCount = Controls.Count
    If controlCount = 0 Then Exit Sub

    ReDim ControlArray(0 To controlCount - 1)

    i = 0
    For Each ctl In Controls
        Set ControlArray(i) = ctl
        i = i + 1
    Next ctl

    SortControlsByTabIndex ControlArray

    For i = LBound(ControlArray) To UBound(ControlArray)
        Set ctl = ControlArray(i)

        WriteControlInfo FormName, ParentName, ctl

        If TypeName(ctl) = "MultiPage" Then
            ExportPages ctl, FormName
        ElseIf HasControls(ctl) Then
            ExportControlsRecursive ctl.Controls, FormName, ctl.Name
        End If
    Next i

End Sub

Private Sub ExportPages(ByVal mp As Object, ByVal FormName As String)

    Dim pg As Object

    For Each pg In mp.Pages
        WriteControlInfo FormName, mp.Name, pg
        ExportControlsRecursive pg.Controls, FormName, pg.Name
    Next pg

End Sub

Private Sub WriteControlInfo( _
        ByVal FormName As String, _
        ByVal ParentName As String, _
        ByVal ctl As Object)

    gStream.WriteText _
        Csv(FormName) & "," & _
        Csv(ParentName) & "," & _
        Csv(ctl.Name) & "," & _
        Csv(TypeName(ctl)) & "," & _
        Csv(GetControlPropertyText(ctl, "Caption")) & "," & _
        Csv(GetControlPropertyText(ctl, "Top")) & "," & _
        Csv(GetControlPropertyText(ctl, "Left")) & "," & _
        Csv(GetControlPropertyText(ctl, "Width")) & "," & _
        Csv(GetControlPropertyText(ctl, "Height")) & "," & _
        Csv(GetControlPropertyText(ctl, "TabIndex")) & _
        vbCrLf

End Sub

Private Function GetControlPropertyText( _
        ByVal ctl As Object, _
        ByVal PropertyName As String) As String

    On Error GoTo PropertyError

    Select Case PropertyName
        Case "Caption"
            GetControlPropertyText = CStr(ctl.Caption)
        Case "Top"
            GetControlPropertyText = CStr(ctl.Top)
        Case "Left"
            GetControlPropertyText = CStr(ctl.Left)
        Case "Width"
            GetControlPropertyText = CStr(ctl.Width)
        Case "Height"
            GetControlPropertyText = CStr(ctl.Height)
        Case "TabIndex"
            GetControlPropertyText = CStr(ctl.TabIndex)
    End Select

    Exit Function

PropertyError:
    Err.Clear
    GetControlPropertyText = ""

End Function

Private Function HasControls(ByVal ctl As Object) As Boolean

    Dim controlCount As Long

    On Error Resume Next
    Err.Clear
    controlCount = ctl.Controls.Count
    HasControls = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

End Function

Private Function Csv(ByVal Value As Variant) As String

    Csv = Chr$(34) & Replace(CStr(Value), Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)

End Function

Private Sub SortControlsByTabIndex(ByRef ControlArray() As Object)

    Dim i As Long
    Dim j As Long
    Dim TempControl As Object

    For i = LBound(ControlArray) To UBound(ControlArray) - 1
        For j = i + 1 To UBound(ControlArray)
            If ControlComesAfter(ControlArray(i), ControlArray(j)) Then
                Set TempControl = ControlArray(i)
                Set ControlArray(i) = ControlArray(j)
                Set ControlArray(j) = TempControl
            End If
        Next j
    Next i

End Sub

Private Function ControlComesAfter( _
        ByVal Control1 As Object, _
        ByVal Control2 As Object) As Boolean

    Dim TabIndex1 As Long
    Dim TabIndex2 As Long
    Dim Top1 As Double
    Dim Top2 As Double
    Dim Left1 As Double
    Dim Left2 As Double

    TabIndex1 = GetControlTabIndex(Control1)
    TabIndex2 = GetControlTabIndex(Control2)

    If TabIndex1 <> TabIndex2 Then
        ControlComesAfter = (TabIndex1 > TabIndex2)
        Exit Function
    End If

    Top1 = GetControlPosition(Control1, "Top")
    Top2 = GetControlPosition(Control2, "Top")

    If Top1 <> Top2 Then
        ControlComesAfter = (Top1 > Top2)
        Exit Function
    End If

    Left1 = GetControlPosition(Control1, "Left")
    Left2 = GetControlPosition(Control2, "Left")

    If Left1 <> Left2 Then
        ControlComesAfter = (Left1 > Left2)
        Exit Function
    End If

    ControlComesAfter = (StrComp(Control1.Name, Control2.Name, vbTextCompare) > 0)

End Function

Private Function GetControlTabIndex(ByVal ctl As Object) As Long

    On Error GoTo NoTabIndex

    GetControlTabIndex = CLng(ctl.TabIndex)
    Exit Function

NoTabIndex:
    GetControlTabIndex = 2147483647
    Err.Clear

End Function

Private Function GetControlPosition( _
        ByVal ctl As Object, _
        ByVal PropertyName As String) As Double

    On Error GoTo PropertyError

    Select Case PropertyName
        Case "Top"
            GetControlPosition = CDbl(ctl.Top)
        Case "Left"
            GetControlPosition = CDbl(ctl.Left)
        Case Else
            GetControlPosition = 0
    End Select

    Exit Function

PropertyError:
    GetControlPosition = 0
    Err.Clear

End Function

