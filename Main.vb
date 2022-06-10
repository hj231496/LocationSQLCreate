
Public Class Main

    Structure ContextMenuStripType
        Dim fruitContextMenuStrip As ContextMenuStrip   'Spread右鍵
        Dim intIndex As Integer
        Dim strName As String
    End Structure



    Public Const PUCEmptyLoc As Integer = 0
    Public Const PUCVirtualLoc As Integer = 8
    Public Const PUCProhibitLoc As Integer = 9
    Public Const PUCOFF As Integer = 0
    Public Const PUCON As Integer = 1
    Public Const PUCDefaultPlant As String = "1480"
    Public Const PUCDefaultWareHouse As String = "E"
    Public Const PUCDefaultBinType As String = "T"

    Dim intSelectedCell(10, 10) As Integer
    Dim intCellColor(10, 10) As Integer


    '三維陣列做出多個Row可設定，每當切換Row，Row(a,b,c)的a跟著切並載入Spread
    '但如果Bay跟Level不同..三維陣列無法這樣宣告..不等長陣列
    '還是一個Row一個Row儲存吧


    '設定 可以設定預設BinType 預設PLANT


    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            ComboBox1.SelectedIndex = 0
            KeyPreview = True
            Object_KeyDown(T_BAY, New KeyEventArgs(Keys.Enter))


        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub
    Private Sub subintCellValueChanged()

        Try
            For i As Integer = 1 To Val(T_BAY.Text)
                For j As Integer = 1 To Val(T_LEVEL.Text)
                    Select Case intCellColor(i, j)
                        Case PUCEmptyLoc
                            S_Location.ActiveSheet.Cells(i, j).BackColor = FL_EmptyLocation_ColorString.BackColor
                        Case PUCVirtualLoc
                            S_Location.ActiveSheet.Cells(i, j).BackColor = FL_Virtual_ColorString.BackColor
                        Case PUCProhibitLoc
                            S_Location.ActiveSheet.Cells(i, j).BackColor = FL_Prohibit_ColorString.BackColor
                    End Select
                Next
            Next

        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub


    Private Sub subContextMenuSet()
        Try
            Dim item As ContextMenuStripType
            item.strName = "A"
            item.fruitContextMenuStrip = New ContextMenuStrip()
            item.fruitContextMenuStrip.Name = "A"
            item.fruitContextMenuStrip.Font = New System.Drawing.Font("Yu Gothic UI", 12.25!, System.Drawing.FontStyle.Bold) 'Spread左邊Row Header大小
            If ComboBox1.SelectedIndex = 0 Then
                item.fruitContextMenuStrip.Items.Add("空庫位")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("禁止庫位")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("虛庫位")
            Else
                item.fruitContextMenuStrip.Items.Add("C:CleanParts")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("P:Pallet")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("I:Hybox")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("T:Tray")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("W:矮籠車")
                item.fruitContextMenuStrip.Items.Add("-")
                item.fruitContextMenuStrip.Items.Add("R:RollCargo")
            End If


            AddHandler item.fruitContextMenuStrip.ItemClicked, AddressOf cms_ItemClick


            S_Location.ContextMenuStrip = item.fruitContextMenuStrip
        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub
    Private Sub Spread_SelectionChanged(sender As Object, e As FarPoint.Win.Spread.SelectionChangedEventArgs) Handles S_Location.SelectionChanged
        Dim sheetView As FarPoint.Win.Spread.SheetView = e.View.GetSheetView()
        Dim cellRange As FarPoint.Win.Spread.Model.CellRange = sheetView.GetSelection(0)
        Dim intFromColumn As Integer
        Dim intFromRow As Integer
        Dim intToColumn As Integer
        Dim intToRow As Integer
        Dim intLoopColumn As Integer
        Dim intLoopRow As Integer


        ReDim intSelectedCell(T_BAY.Text, T_LEVEL.Text)
        intFromColumn = cellRange.Column
        intFromRow = cellRange.Row
        intToColumn = cellRange.Column + cellRange.ColumnCount - 1
        intToRow = cellRange.Row + cellRange.RowCount - 1

        For intLoopColumn = intFromColumn To intToColumn
            For intLoopRow = intFromRow To intToRow
                If intFromRow <> 0 AndAlso intToRow <> 0 AndAlso intFromColumn <> 0 AndAlso intFromRow <> 0 Then

                    intSelectedCell(intLoopRow, intLoopColumn) = 1

                End If
            Next
        Next
    End Sub
    Private Sub cms_ItemClick(sender As Object, e As ToolStripItemClickedEventArgs)
        Try
            '找到圈起來的那些
            '根據點選項目，更新顏色跟intCell
            For i As Integer = 0 To T_BAY.Text
                For j As Integer = 0 To T_LEVEL.Text
                    If intSelectedCell(i, j) = 1 Then
                        If ComboBox1.SelectedIndex = 0 Then
                            Select Case e.ClickedItem.Text.ToString
                                Case FL_EmptyLocation.Text
                                    intCellColor(i, j) = PUCEmptyLoc
                                Case FL_Prohibit.Text
                                    intCellColor(i, j) = PUCProhibitLoc
                                Case FL_Virtual.Text
                                    intCellColor(i, j) = PUCVirtualLoc
                            End Select
                        Else
                            S_Location.ActiveSheet.Cells(i, j).Value = fnComboDataGet(e.ClickedItem.Text.ToString)
                        End If
                    End If
                Next
            Next


            '值改完要重新遍歷Show Color
            subintCellValueChanged()

        Catch ex As Exception
            logger.Error(ex.Message)
        End Try

    End Sub
    Private Sub T_BAY_KeyPress(sender As Object, e As KeyPressEventArgs) Handles T_ROW.KeyPress, T_BAY.KeyPress, T_LEVEL.KeyPress
        Try
            If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) And Asc(e.KeyChar) <> 8 Then
                e.Handled = True
            End If
        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub
    Private Sub Object_KeyDown(sender As Object, e As KeyEventArgs) Handles T_BAY.KeyDown, T_LEVEL.KeyDown
        Try
            If (T_BAY.Text.Trim <> "0" And T_BAY.Text.Trim <> "") And (T_LEVEL.Text.Trim <> "0" And T_LEVEL.Text.Trim <> "") Then

                ReDim intCellColor(T_BAY.Text, T_LEVEL.Text) '要保留嗎

                subSetSpreadSize(S_Location, T_BAY.Text, T_LEVEL.Text)

                For i As Integer = 1 To S_Location.ActiveSheet.RowCount - 1
                    For j As Integer = 1 To S_Location.ActiveSheet.ColumnCount - 1

                        Call subSetSpreadFont(S_Location, i, j, 1, 1)

                        S_Location.ActiveSheet.Cells(i, j).Row.Height = 40
                        S_Location.ActiveSheet.Cells(i, j).Column.Width = 40

                        S_Location.ActiveSheet.Cells(i, j).Text = PUCDefaultBinType
                        S_Location.ActiveSheet.Cells(i, j).BackColor = Color.White
                    Next
                Next
            End If

        Catch ex As Exception
            logger.Error(ex.Message)
        End Try

    End Sub
    Private Sub subSetSpreadSize(ByRef objSpread As FarPoint.Win.Spread.FpSpread, ByVal intMaxBay As Integer, ByVal intMaxLevel As Integer)
        Try

            objSpread.ActiveSheet.ColumnCount = intMaxBay + 1 '設定Spread的長度(Bay)
            objSpread.ActiveSheet.RowCount = intMaxLevel + 1 '設定Spread的高度(Level)

            'Set Bay Title No
            For lonCol = 0 To intMaxBay - 1
                Call subSetSpreadFont(objSpread, 0, lonCol + 1, 0, 0)
                objSpread.ActiveSheet.Cells(0, lonCol + 1).Value = lonCol + 1
            Next

            'Set Level Title No
            For lonRow = 0 To intMaxLevel - 1
                Call subSetSpreadFont(objSpread, lonRow + 1, 0, 0, 0)
                objSpread.ActiveSheet.Cells(lonRow + 1, 0).Value = intMaxLevel - lonRow
            Next

            If intMaxBay > 27 AndAlso intMaxLevel > 9 Then
                '會產生SScrBar  17MM
                objSpread.Width = 45 + 27 * 40 + 17
                objSpread.Height = 45 + 9 * 40 + 17
            ElseIf intMaxBay <= 27 AndAlso intMaxLevel <= 9 Then
                objSpread.Width = 45 + intMaxBay * 40
                objSpread.Height = 45 + intMaxLevel * 40
            ElseIf intMaxBay <= 27 AndAlso intMaxLevel > 9 Then
                objSpread.Width = 45 + intMaxBay * 40 + 17
                objSpread.Height = 45 + 9 * 40
            ElseIf intMaxLevel <= 9 AndAlso intMaxBay > 27 Then
                objSpread.Height = 45 + intMaxLevel * 40 + 17
                objSpread.Width = 45 + 27 * 40
            End If


            If objSpread.Width + 80 >= 880 Then
                Me.Width = objSpread.Width + 80
            End If

        Catch ex As Exception
            logger.Error(ex.Message)
        End Try

    End Sub
    Public Sub subSetSpreadFont(ByRef objSpread As FarPoint.Win.Spread.FpSpread, ByVal intSetRow As Integer, ByVal intSetCol As Integer,
                                  ByVal intStaRow As Integer, ByVal intStaCol As Integer)

        objSpread.ActiveSheet.Cells(intSetRow, intSetCol).BackColor = objSpread.ActiveSheet.Cells(intStaRow, intStaCol).BackColor
        objSpread.ActiveSheet.Cells(intSetRow, intSetCol).VerticalAlignment = objSpread.ActiveSheet.Cells(intStaRow, intStaCol).VerticalAlignment
        objSpread.ActiveSheet.Cells(intSetRow, intSetCol).HorizontalAlignment = objSpread.ActiveSheet.Cells(intStaRow, intStaCol).HorizontalAlignment
        objSpread.ActiveSheet.Cells(intSetRow, intSetCol).Border = objSpread.ActiveSheet.Cells(intStaRow, intStaCol).Border
        objSpread.ActiveSheet.Cells(intSetRow, intSetCol).CellType = objSpread.ActiveSheet.Cells(intStaRow, intStaCol).CellType
    End Sub
    Private Sub S_Location_CellDoubleClick(sender As Object, e As FarPoint.Win.Spread.CellClickEventArgs) Handles S_Location.CellDoubleClick
        Try
            If ComboBox1.SelectedIndex = 0 Then
                If e.Row >= 1 And e.Row < S_Location.ActiveSheet.RowCount And
                   e.Column >= 1 And e.Row < S_Location.ActiveSheet.ColumnCount Then
                    Select Case intCellColor(e.Row, e.Column)
                        Case PUCEmptyLoc '空庫位->虛庫位                     
                            intCellColor(e.Row, e.Column) = PUCVirtualLoc
                        Case PUCVirtualLoc '虛庫位->禁止庫位
                            intCellColor(e.Row, e.Column) = PUCProhibitLoc
                        Case PUCProhibitLoc '禁止庫位->空庫位
                            intCellColor(e.Row, e.Column) = PUCEmptyLoc
                    End Select
                End If
            End If


            '值改完要重新遍歷Show Color
            subintCellValueChanged()
        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            subContextMenuSet()

            If ComboBox1.SelectedIndex = 0 Then
                For i As Integer = 1 To S_Location.ActiveSheet.RowCount - 1
                    For j As Integer = 1 To S_Location.ActiveSheet.ColumnCount - 1
                        S_Location.ActiveSheet.Cells(i, j).Locked = True
                    Next
                Next
            Else
                For i As Integer = 1 To S_Location.ActiveSheet.RowCount - 1
                    For j As Integer = 1 To S_Location.ActiveSheet.ColumnCount - 1
                        S_Location.ActiveSheet.Cells(i, j).Locked = False
                    Next
                Next
            End If

        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub

    Private Sub Btn_Create_Click(sender As Object, e As EventArgs) Handles Btn_Create.Click
        Dim strSQL As String
        Dim strFullLocation As String
        Dim strCrane As String = Math.Ceiling(T_ROW.Text / 2)
        'Range (1,1)~(20,20)，但她是倒過來的

        Try


            If R_ORACLE.Checked = True Then
                strSQL = "INSERT ALL " & vbCrLf

                For i As Integer = 1 To Val(T_BAY.Text)
                    For j As Integer = 1 To Val(T_BAY.Text)
                        '倒過來 所以用最大扣掉i然後+1
                        strFullLocation = T_ROW.Text.PadLeft(2, "0") & (Val(T_BAY.Text) - i + 1).ToString.PadLeft(3, "0") & (Val(T_LEVEL.Text) - j + 1).ToString.PadLeft(2, "0")

                        strSQL &= "INTO LOCATION VALUES ( "
                        strSQL &= fnSQLStrSet(PUCDefaultPlant) & " , " 'PLANT
                        strSQL &= fnSQLStrSet(strCrane) & " , " 'CRANE
                        strSQL &= fnSQLStrSet(strFullLocation) & " , " 'LOCATION
                        strSQL &= fnSQLStrSet(S_Location.ActiveSheet.Cells(i, j).Value) & " , " 'BINTYPE
                        strSQL &= fnSQLStrSet("0") & " , " 'SEQ
                        If intCellColor(i, j) <> PUCProhibitLoc Then
                            strSQL &= fnSQLStrSet(intCellColor(i, j)) & " , " 'STAT
                            strSQL &= fnSQLStrSet(PUCOFF) & " , " 'PROHIBIT
                        Else
                            strSQL &= fnSQLStrSet(PUCEmptyLoc) & " , " 'STAT
                            strSQL &= fnSQLStrSet(PUCON) & " , " 'PROHIBIT
                        End If

                        strSQL &= fnSQLStrSet(" ") & " , " 'PROPERTY
                        strSQL &= "SYSDATE" & " , " 'TIMESTAMP
                        strSQL &= "NULL"  'PALLET

                        strSQL &= " ) " & vbCrLf
                    Next
                Next

                strSQL &= "SELECT * FROM DUAL "

            End If



            '用個小視窗顯示產生結果，若要繼續地話儲存至
            Dim Form As New Result
            Form.TextBox1.Text = strSQL
            Form.ShowDialog()


        Catch ex As Exception
            logger.Error(ex.Message)
        End Try

    End Sub

    Private Sub S_Location_KeyDown(sender As Object, e As KeyEventArgs) Handles S_Location.KeyDown
        Try
            If ComboBox1.SelectedIndex = 1 Then
                If S_Location.ActiveSheet.ActiveRowIndex >= 1 And S_Location.ActiveSheet.ActiveRowIndex < S_Location.ActiveSheet.RowCount And
                   S_Location.ActiveSheet.ActiveColumnIndex >= 1 And S_Location.ActiveSheet.ActiveColumnIndex < S_Location.ActiveSheet.ColumnCount Then
                    S_Location.ActiveSheet.ActiveCell.Value = ""
                End If
            End If
        Catch ex As Exception
            logger.Error(ex.Message)
        End Try
    End Sub

    Private Sub Setting_Click(sender As Object, e As EventArgs) Handles Setting.Click
        Try




        Catch ex As Exception
            logger.Error(ex.Message)
        End Try

    End Sub

    Public Function fnSQLStrSet(ByVal strText As String) As String

        If strText Is Nothing Then
            fnSQLStrSet = "' '"
        Else

            fnSQLStrSet = "'" & strText.Replace("'", "''") & "'"
        End If

    End Function
    Public Function fnComboDataGet(ByRef strText As String) As String
        Dim intPos As Short
        Dim strGetText As String

        intPos = InStr(1, strText, ":")

        If intPos < 1 Then
            strGetText = strText
        Else
            strGetText = Microsoft.VisualBasic.Left(strText, intPos - 1)
        End If

        fnComboDataGet = strGetText
    End Function
End Class


Module NLog7
    '*-----------------------------------------------------*
    '*  NLog用
    '*-----------------------------------------------------*
    Public Const loggerName = "Process"
    Public logger As NLog.Logger = NLog.LogManager.GetLogger(loggerName)

End Module
