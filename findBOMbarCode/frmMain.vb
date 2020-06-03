Imports System.Net
Imports findBOMbarCode.ModbusTCP

Public Class frmMain
    Inherits System.Windows.Forms.Form
    Private MBMaster As ModbusTCP.Master
    Private txtdata1, txtdata2, txtdata3, txtdata4, txtdata5, txtdata6 As TextBox
    Private labData1, labData2, labData3, labData4, labData5, labData6 As Label
    Private data1, data2, data3, data4, data5, data6 As Byte()

    Dim txtIP As TextBox = New TextBox
    Dim txtStartAdress As TextBox = New TextBox
    Dim txtUnit As TextBox = New TextBox
    Dim txtSize As TextBox = New TextBox

    Dim word1 As Integer() = New Integer(0) {}
    Dim word2 As Integer() = New Integer(0) {}
    Dim word3 As Integer() = New Integer(0) {}
    Dim word4 As Integer() = New Integer(0) {}
    Dim word5 As Integer() = New Integer(0) {}
    Dim word6 As Integer() = New Integer(0) {}

    Dim words1 As Integer() = New Integer(0.0) {}
    Dim words2 As Integer() = New Integer(0) {}
    Dim words3 As Integer() = New Integer(0) {}
    Dim words4 As Integer() = New Integer(0) {}
    Dim words5 As Integer() = New Integer(0) {}
    Dim words6 As Integer() = New Integer(0) {}

    Dim wordx1 As Double() = New Double(0.0) {}
    Dim wordx2 As Double() = New Double(0.0) {}
    Dim wordx3 As Double() = New Double(0.0) {}
    Dim wordx4 As Double() = New Double(0.0) {}
    Dim wordx5 As Double() = New Double(0.0) {}
    Dim wordx6 As Double() = New Double(0.0) {}

    Dim wordxx As Integer = 65536

    Dim hexCheck1, hexCheck2, hexCheck3, hexCheck4, hexCheck5, hexCheck6, hexCheck7, hexCheck8, hexCheck9, hexCheck10, hexCheck58, hexCheck60 As New TextBox
    Dim txtBarcode, txtBarcode1, txtBarcode2, txtBarcode3, txtBarcode4, txtBarcode5, txtBarcode6 As New TextBox
    Dim stScale1, stScale2, stScale3, stScale4, stScale5, stScale6 As Integer
    Dim stDoScale1, stDoScale2, stDoScale3, stDoScale4, stDoScale5, stDoScale6 As Integer
    Dim chkFScale1, chkFScale2, chkFScale3, chkFScale4, chkFScale5, chkFScale6 As Integer
    Dim chkSaveScale1, chkSaveScale2, chkSaveScale3, chkSaveScale4, chkSaveScale5, chkSaveScale6 As Integer
    Dim nrec, nrec1, nrec2, nrec3, nrec4, nrec5, nrec6 As Integer
    Dim coub1, coub2, coub3, coub4, coub5, coub6 As Integer
    Dim nc1, nc2, nc3, nc4, nc5, nc6 As Integer
    Dim cstart1, cstart2, cstart3, cstart4, cstart5, cstart6 As Integer
    Dim sendWeight1, sendWeight2, sendWeight3, sendWeight4, sendWeight5, sendWeight6 As Integer
    Dim recWeight1, recWeight2, recWeight3, recWeight4, recWeight5, recWeight6 As Decimal
    Dim codeS1v1, codeS1V2, codeS1V3, codeS1V4, codeS1V5, codeS1V6 As String
    Dim codeS2v1, codeS2V2, codeS2V3, codeS2V4, codeS2V5, codeS2V6 As String
    Dim codeS3v1, codeS3V2, codeS3V3, codeS3V4, codeS3V5, codeS3V6 As String
    Dim codeS4v1, codeS4V2, codeS4V3, codeS4V4, codeS4V5, codeS4V6 As String
    Dim codeS5v1, codeS5V2, codeS5V3, codeS5V4, codeS5V5, codeS5V6 As String
    Dim codeS6v1, codeS6V2, codeS6V3, codeS6V4, codeS6V5, codeS6V6 As String
    Dim activeScale1, activeScale2, activeScale3, activeScale4, activeScale5, activeScale6 As Integer
    Dim sqlStr1, sqlStr2, sqlStr3, sqlStr4, sqlStr5, sqlStr6 As String
    Dim activeBarScale1, activeBarScale2, activeBarScale3, activeBarScale4, activeBarScale5, activeBarScale6 As String
    Dim sqllist As String = ""

    Dim lbStkName, lbPCName, lbQty, lbQtyW, lbTotalQtyW, lbTankQty, lbTankName As New Label
    Dim txtStkCode, txtPCcode As New TextBox
    Dim TankID As String = ""
    Dim TankNum As Integer = 0
    Dim TankName As String = ""
    Dim TankW As Double = 0

    Sub lsvScaleformat(numScale As Integer)
        Select Case numScale
            Case 1
                lsvScale1.Columns.Add("#", 30, HorizontalAlignment.Right) '1====== 0   
                lsvScale1.Columns.Add("รหัส", 80, HorizontalAlignment.Left) '1====== 2     
                lsvScale1.Columns.Add("ชื่อวัตถุดิบ", 180, HorizontalAlignment.Left) '1====== 1               
                lsvScale1.Columns.Add("ปริมาณ", 150, HorizontalAlignment.Right) '1====== 3 
                lsvScale1.Columns.Add("หน่วย", 60, HorizontalAlignment.Left) '1====== 3 
                lsvScale1.Columns.Add("วาล์วที่", 60, HorizontalAlignment.Center) '1====== 3 
                lsvScale1.View = View.Details
                lsvScale1.GridLines = True
                Exit Select
            Case 2
                lsvScale2.Columns.Add("#", 30, HorizontalAlignment.Right) '1====== 0   
                lsvScale2.Columns.Add("รหัส", 80, HorizontalAlignment.Left) '1====== 2     
                lsvScale2.Columns.Add("ชื่อวัตถุดิบ", 180, HorizontalAlignment.Left) '1====== 1               
                lsvScale2.Columns.Add("ปริมาณ", 150, HorizontalAlignment.Right) '1====== 3 
                lsvScale2.Columns.Add("หน่วย", 60, HorizontalAlignment.Left) '1====== 3 
                lsvScale2.Columns.Add("วาล์วที่", 60, HorizontalAlignment.Center) '1====== 3 
                lsvScale2.View = View.Details
                lsvScale2.GridLines = True
                Exit Select
            Case 3
                lsvScale3.Columns.Add("#", 30, HorizontalAlignment.Right) '1====== 0   
                lsvScale3.Columns.Add("รหัส", 80, HorizontalAlignment.Left) '1====== 2     
                lsvScale3.Columns.Add("ชื่อวัตถุดิบ", 180, HorizontalAlignment.Left) '1====== 1               
                lsvScale3.Columns.Add("ปริมาณ", 150, HorizontalAlignment.Right) '1====== 3 
                lsvScale3.Columns.Add("หน่วย", 60, HorizontalAlignment.Left) '1====== 3 
                lsvScale3.Columns.Add("วาล์วที่", 60, HorizontalAlignment.Center) '1====== 3 
                lsvScale3.View = View.Details
                lsvScale3.GridLines = True
                Exit Select
            Case 4
                lsvScale4.Columns.Add("#", 30, HorizontalAlignment.Right) '1====== 0   
                lsvScale4.Columns.Add("รหัส", 80, HorizontalAlignment.Left) '1====== 2     
                lsvScale4.Columns.Add("ชื่อวัตถุดิบ", 180, HorizontalAlignment.Left) '1====== 1               
                lsvScale4.Columns.Add("ปริมาณ", 150, HorizontalAlignment.Right) '1====== 3 
                lsvScale4.Columns.Add("หน่วย", 60, HorizontalAlignment.Left) '1====== 3 
                lsvScale4.Columns.Add("วาล์วที่", 60, HorizontalAlignment.Center) '1====== 3 
                lsvScale4.View = View.Details
                lsvScale4.GridLines = True
                Exit Select
            Case 5
                lsvScale5.Columns.Add("#", 30, HorizontalAlignment.Right) '1====== 0   
                lsvScale5.Columns.Add("รหัส", 80, HorizontalAlignment.Left) '1====== 2     
                lsvScale5.Columns.Add("ชื่อวัตถุดิบ", 180, HorizontalAlignment.Left) '1====== 1               
                lsvScale5.Columns.Add("ปริมาณ", 150, HorizontalAlignment.Right) '1====== 3 
                lsvScale5.Columns.Add("หน่วย", 60, HorizontalAlignment.Left) '1====== 3 
                lsvScale5.Columns.Add("วาล์วที่", 60, HorizontalAlignment.Center) '1====== 3 
                lsvScale5.View = View.Details
                lsvScale5.GridLines = True
                Exit Select
            Case 6
                lsvScale6.Columns.Add("#", 30, HorizontalAlignment.Right) '1====== 0   
                lsvScale6.Columns.Add("รหัส", 80, HorizontalAlignment.Left) '1====== 2     
                lsvScale6.Columns.Add("ชื่อวัตถุดิบ", 180, HorizontalAlignment.Left) '1====== 1               
                lsvScale6.Columns.Add("ปริมาณ", 150, HorizontalAlignment.Right) '1====== 3 
                lsvScale6.Columns.Add("หน่วย", 60, HorizontalAlignment.Left) '1====== 3 
                lsvScale6.Columns.Add("วาล์วที่", 60, HorizontalAlignment.Center) '1====== 3 
                lsvScale6.View = View.Details
                lsvScale6.GridLines = True
                Exit Select
        End Select
    End Sub

    Function getQtyW(stkCode As String) As Double
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        txtSQL = "Select * "
        txtSQL = txtSQL & "From BOMmast "
        txtSQL = txtSQL & "Where BOM_Stk_Code='" & stkCode & "' "
        txtSQL = txtSQL & "And BOM_RM_Code='04003' "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "master")
        Return subDS.Tables("master").Rows(0).Item("BOM_RM_Values")

    End Function

    Private Sub getTank(qtyW As Double)

        Dim numberP As Integer = 0
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        Dim iRow As Integer = 0
        Dim nCount As Integer = 0
        Dim chkAns As Integer = 0
        Dim tank As Double = 0
        Dim nTank As Integer = 0
        txtSQL = "Select * "
        txtSQL = txtSQL & "From BOM_RMmast "
        txtSQL = txtSQL & "Where BOM_GRp_ID='08' "
        txtSQL = txtSQL & "Order by BOM_Dtl_Se "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "master")
        nCount = subDS.Tables("master").Rows.Count
        iRow = 0
        nTank = 1
        Do Until chkAns = 1
            TankID = subDS.Tables("master").Rows(iRow).Item("rm_Code")
            TankNum = nTank 'iRow + 1
            TankName = subDS.Tables("master").Rows(iRow).Item("BOM_Dtl_Name")
            'TankName = CDbl(subDS.Tables("master").Rows(iRow).Item("BOM_Dtl_Unit_1"))
            tank = subDS.Tables("master").Rows(iRow).Item("BOM_Dtl_Unit_1")
            TankW = (subDS.Tables("master").Rows(iRow).Item("BOM_Dtl_Unit_1"))

            If qtyW > (tank * (nTank)) Then
                If nCount - 1 = iRow Then
                    iRow = 0
                    nTank = nTank + 1
                Else
                    iRow = iRow + 1
                End If
            Else
                chkAns = 1
            End If
        Loop

    End Sub

    Private Function ConvertHexToDec(ByVal ckhex As Integer)
        Dim returnWord As String = ""
        Select Case ckhex
            Case 0
                returnWord = ""
                Exit Select
            Case 40
                returnWord = "("
                Exit Select
            Case 41
                returnWord = ")"
                Exit Select
            Case 42
                returnWord = "*"
                Exit Select
            Case 43
                returnWord = "+"
                Exit Select
            Case 44
                returnWord = ","
                Exit Select
            Case 45
                returnWord = "-"
                Exit Select
            Case 46
                returnWord = "."
                Exit Select
            Case 47
                returnWord = "/"
                Exit Select
            Case 48
                returnWord = "0"
                Exit Select
            Case 49
                returnWord = "1"
                Exit Select
            Case 50
                returnWord = "2"
                Exit Select
            Case 51
                returnWord = "3"
                Exit Select
            Case 52
                returnWord = "4"
                Exit Select
            Case 53
                returnWord = "5"
                Exit Select
            Case 54
                returnWord = "6"
                Exit Select
            Case 55
                returnWord = "7"
                Exit Select
            Case 56
                returnWord = "8"
                Exit Select
            Case 57
                returnWord = "9"
                Exit Select
            Case 58
                returnWord = ":"
                Exit Select
            Case 59
                returnWord = ";"
                Exit Select
            Case 60
                returnWord = "<"
                Exit Select
            Case 61
                returnWord = "="
                Exit Select
            Case 62
                returnWord = ">"
                Exit Select
            Case 63
                returnWord = "?"
                Exit Select
            Case 64
                returnWord = "@"
                Exit Select
            Case 65
                returnWord = "A"
                Exit Select
            Case 66
                returnWord = "B"
                Exit Select
            Case 67
                returnWord = "C"
                Exit Select
            Case 68
                returnWord = "D"
                Exit Select
            Case 69
                returnWord = "E"
                Exit Select
            Case 70
                returnWord = "F"
                Exit Select
            Case 71
                returnWord = "G"
                Exit Select
            Case 72
                returnWord = "H"
                Exit Select
            Case 73
                returnWord = "I"
                Exit Select
            Case 74
                returnWord = "J"
                Exit Select
            Case 75
                returnWord = "K"
                Exit Select
            Case 76
                returnWord = "L"
                Exit Select
            Case 77
                returnWord = "M"
                Exit Select
            Case 78
                returnWord = "N"
                Exit Select
            Case 79
                returnWord = "O"
                Exit Select
            Case 80
                returnWord = "P"
                Exit Select
            Case 81
                returnWord = "Q"
                Exit Select
            Case 82
                returnWord = "R"
                Exit Select
            Case 83
                returnWord = "S"
                Exit Select
            Case 84
                returnWord = "T"
                Exit Select
            Case 85
                returnWord = "U"
                Exit Select
            Case 86
                returnWord = "V"
                Exit Select
            Case 87
                returnWord = "W"
                Exit Select
            Case 88
                returnWord = "X"
                Exit Select
            Case 89
                returnWord = "Y"
                Exit Select
            Case 90
                returnWord = "Z"
                Exit Select
            Case 91
                returnWord = "["
                Exit Select
            Case 92
                returnWord = "\"
                Exit Select
            Case 93
                returnWord = "]"
                Exit Select
            Case 94
                returnWord = "^"
                Exit Select
            Case 95
                returnWord = "_"
                Exit Select
            Case 96
                returnWord = "'"
                Exit Select
            Case 97
                returnWord = "a"
                Exit Select
            Case 98
                returnWord = "b"
                Exit Select
            Case 99
                returnWord = "c"
                Exit Select
            Case 100
                returnWord = "d"
                Exit Select
            Case 101
                returnWord = "e"
                Exit Select
            Case 102
                returnWord = "f"
                Exit Select
            Case 103
                returnWord = "g"
                Exit Select
            Case 104
                returnWord = "h"
                Exit Select
            Case 105
                returnWord = "i"
                Exit Select
            Case 106
                returnWord = "j"
                Exit Select
            Case 107
                returnWord = "k"
                Exit Select
            Case 108
                returnWord = "l"
                Exit Select
            Case 109
                returnWord = "m"
                Exit Select
            Case 110
                returnWord = "n"
                Exit Select
            Case 111
                returnWord = "o"
                Exit Select
            Case 112
                returnWord = "p"
                Exit Select
            Case 113
                returnWord = "q"
                Exit Select
            Case 114
                returnWord = "r"
                Exit Select
            Case 115
                returnWord = "s"
                Exit Select
            Case 116
                returnWord = "t"
                Exit Select
            Case 117
                returnWord = "u"
                Exit Select
            Case 118
                returnWord = "v"
                Exit Select
            Case 119
                returnWord = "w"
                Exit Select
            Case 120
                returnWord = "x"
                Exit Select
            Case 121
                returnWord = "y"
                Exit Select
            Case 122
                returnWord = "z"
                Exit Select
            Case 123
                returnWord = "{"
                Exit Select
            Case 124
                returnWord = "|"
                Exit Select
            Case 125
                returnWord = "}"
                Exit Select
            Case 126
                returnWord = "~"
                Exit Select
        End Select

        Return returnWord
    End Function

    Private Function ConvertDecToHex(ByVal ckDec As String)
        Dim returnHex As Integer
        Select Case ckDec
            Case ""
                returnHex = 0
                Exit Select
            Case "("
                returnHex = 40
                Exit Select
            Case ")"
                returnHex = 41
                Exit Select
            Case "*"
                returnHex = 42
                Exit Select
            Case "+"
                returnHex = 43
                Exit Select
            Case ","
                returnHex = 44
                Exit Select
            Case "-"
                returnHex = 45
                Exit Select
            Case "."
                returnHex = 46
                Exit Select
            Case "/"
                returnHex = 47
                Exit Select
            Case "0"
                returnHex = 48
                Exit Select
            Case "1"
                returnHex = 49
                Exit Select
            Case "2"
                returnHex = 50
                Exit Select
            Case "3"
                returnHex = 51
                Exit Select
            Case "4"
                returnHex = 52
                Exit Select
            Case "5"
                returnHex = 53
                Exit Select
            Case "6"
                returnHex = 54
                Exit Select
            Case "7"
                returnHex = 55
                Exit Select
            Case "8"
                returnHex = 56
                Exit Select
            Case "9"
                returnHex = 57
                Exit Select
            Case ":"
                returnHex = 58
                Exit Select
            Case ";"
                returnHex = 59
                Exit Select
            Case "<"
                returnHex = 60
                Exit Select
            Case "="
                returnHex = 61
                Exit Select
            Case ">"
                returnHex = 62
                Exit Select
            Case "?"
                returnHex = 63
                Exit Select
            Case "@"
                returnHex = 64
                Exit Select
            Case "A"
                returnHex = 65
                Exit Select
            Case "B"
                returnHex = 66
                Exit Select
            Case "C"
                returnHex = 67
                Exit Select
            Case "D"
                returnHex = 68
                Exit Select
            Case "E"
                returnHex = 69
                Exit Select
            Case "F"
                returnHex = 70
                Exit Select
            Case "G"
                returnHex = 71
                Exit Select
            Case "H"
                returnHex = 72
                Exit Select
            Case "I"
                returnHex = 73
                Exit Select
            Case "J"
                returnHex = 74
                Exit Select
            Case "K"
                returnHex = 75
                Exit Select
            Case "L"
                returnHex = 76
                Exit Select
            Case "M"
                returnHex = 77
                Exit Select
            Case "N"
                returnHex = 78
                Exit Select
            Case "O"
                returnHex = 79
                Exit Select
            Case "P"
                returnHex = 80
                Exit Select
            Case "Q"
                returnHex = 81
                Exit Select
            Case "R"
                returnHex = 82
                Exit Select
            Case "S"
                returnHex = 83
                Exit Select
            Case "T"
                returnHex = 84
                Exit Select
            Case "U"
                returnHex = 85
                Exit Select
            Case "V"
                returnHex = 86
                Exit Select
            Case "W"
                returnHex = 87
                Exit Select
            Case "X"
                returnHex = 88
                Exit Select
            Case "Y"
                returnHex = 89
                Exit Select
            Case "Z"
                returnHex = 90
                Exit Select
            Case "["
                returnHex = 91
                Exit Select
            Case "\"
                returnHex = 92
                Exit Select
            Case "]"
                returnHex = 93
                Exit Select
            Case "^"
                returnHex = 94
                Exit Select
            Case "_"
                returnHex = 95
                Exit Select
            Case "'"
                returnHex = 96
                Exit Select
            Case "a"
                returnHex = 97
                Exit Select
            Case "b"
                returnHex = 98
                Exit Select
            Case "c"
                returnHex = 99
                Exit Select
            Case "d"
                returnHex = 100
                Exit Select
            Case "e"
                returnHex = 101
                Exit Select
            Case "f"
                returnHex = 102
                Exit Select
            Case "g"
                returnHex = 103
                Exit Select
            Case "h"
                returnHex = 104
                Exit Select
            Case "i"
                returnHex = 105
                Exit Select
            Case "j"
                returnHex = 106
                Exit Select
            Case "k"
                returnHex = 107
                Exit Select
            Case "l"
                returnHex = 108
                Exit Select
            Case "m"
                returnHex = 109
                Exit Select
            Case "n"
                returnHex = 110
                Exit Select
            Case "o"
                returnHex = 111
                Exit Select
            Case "p"
                returnHex = 112
                Exit Select
            Case "q"
                returnHex = 113
                Exit Select
            Case "r"
                returnHex = 114
                Exit Select
            Case "s"
                returnHex = 115
                Exit Select
            Case "t"
                returnHex = 116
                Exit Select
            Case "u"
                returnHex = 117
                Exit Select
            Case "v"
                returnHex = 118
                Exit Select
            Case "w"
                returnHex = 119
                Exit Select
            Case "x"
                returnHex = 120
                Exit Select
            Case "y"
                returnHex = 121
                Exit Select
            Case "z"
                returnHex = 122
                Exit Select
            Case "{"
                returnHex = 123
                Exit Select
            Case "|"
                returnHex = 124
                Exit Select
            Case "}"
                returnHex = 125
                Exit Select
            Case "~"
                returnHex = 126
                Exit Select
        End Select

        Return returnHex
    End Function
    Function getDocNo5(strDocNo As String, numSCale As Integer) As String
        Dim Ans As String = ""

        If strDocNo = "35819AB" And numSCale = 5 Then
            ' Dim strDocNo As String = ""
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataH "
            txtSQL = txtSQL & "Where TranDataH.Trh_Type='M'  "
            txtSQL = txtSQL & "And TranDataH.Trh_Chk_Print='0' "
            txtSQL = txtSQL & "Order by TranDataH.Trh_Date desc "

            Dim subDS5 As New DataSet
            Dim subDA5 As SqlClient.SqlDataAdapter

            subDA5 = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA5.Fill(subDS5, "docList")

            If subDS5.Tables("docList").Rows.Count > 0 Then
                Ans = subDS5.Tables("docList").Rows(0).Item("Trh_No")
            End If
            subDS5 = Nothing
            subDA5 = Nothing
        Else
            Ans = strDocNo
        End If

        Return Ans

    End Function
    Sub showData(strDocNo As String, numScale As Integer)
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        Dim lvi As ListViewItem
        Dim anyData() As String
        Dim i As Integer = 0
        Dim chkData As Boolean = False
        Dim chkRow As Integer = 0
        Dim dblTotal As Double = 0

        txtSQL = "Select * "
        txtSQL = txtSQL & "From BOM_ValveSet "
        txtSQL = txtSQL & "Left Join BOM_RMmast "
        txtSQL = txtSQL & "On BOM_ValveSet.Val_RM_Code=BOM_RMmast.rm_Code "
        txtSQL = txtSQL & "Order by Val_RM_Number "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "master0")
        'lsvBOM01.Clear()
        'lsvBOMformat()
        If numScale = 1 Then lsvScale1.Clear()
        If numScale = 2 Then lsvScale2.Clear()
        If numScale = 3 Then lsvScale3.Clear()
        If numScale = 4 Then lsvScale4.Clear()
        If numScale = 5 Then lsvScale5.Clear()
        If numScale = 6 Then lsvScale6.Clear()
        lsvScaleformat(numScale)

        For n = 0 To subDS.Tables("master0").Rows.Count - 1
            Dim rmName As String = ""
            Dim rmCode As String = ""
            Dim rmValues As Double = 0
            Dim rmUnit As String = ""
            Dim rmValve As String = ""

            strDocNo = getDocNo5(strDocNo, numScale)  '  เพิ่มเพื่อเปลี่ยนเลขที่ใบคุม  18-02-2562  Kritpon  

            txtSQL = "Select * "
            txtSQL = txtSQL & "From BOMmastH "
            txtSQL = txtSQL & "left Join BOMmastD "
            txtSQL = txtSQL & "On BOMmastH.BOM_No=BOMmastD.BOM_No "
            txtSQL = txtSQL & "left Join BOM_RMmast "
            txtSQL = txtSQL & "On BomMastD.BOM_RM_Code=BOM_RMmast.rm_Code "
            txtSQL = txtSQL & "Left Join BaseMast "
            txtSQL = txtSQL & "On BOM_Stk_Code=BaseMast.Stk_Code "

            txtSQL = txtSQL & "Left Join TranDataD_E "
            txtSQL = txtSQL & "On BOMmastH.TRh_No=TranDataD_E.Dtl_No "

            txtSQL = txtSQL & "Where BOMmastH.BOM_No='" & strDocNo & "' "
            txtSQL = txtSQL & "And BOMmastD.BOM_RM_Code = '" & subDS.Tables("master0").Rows(n).Item("Val_RM_Code") & "' "
            txtSQL = txtSQL & "Order by  BOM_RM_Code "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            If chkData = True Then
                subDS.Tables("master").Clear()
                chkData = False
            End If
            subDA.Fill(subDS, "master")
            chkData = True

            If subDS.Tables("master").Rows.Count > 0 Then

                With subDS.Tables("master").Rows(0)

                    lbStkName.Text = .Item("Stk_Name_1")
                    txtStkCode.Text = .Item("Stk_Code")
                    lbPCName.Text = .Item("Stk_PC_Name")

                    txtPCcode.Text = .Item("Stk_Code_PC")

                    lbQty.Text = .Item("BOM_Qty")
                    lbQtyW.Text = Format(getQtyW(txtStkCode.Text), "#,##0.0")
                    lbTotalQtyW.Text = Format((lbQty.Text * lbQtyW.Text), "#,##0.0")
                    Call getTank(lbTotalQtyW.Text)
                    lbTankQty.Text = TankNum
                    lbTankName.Text = TankW '.Item("BOM_Dtl_Unit_1")  ' BOM_Dtl_Unit_1

                    Select Case numScale
                        Case 1
                            words1(84) = ConvertDecToHex(Mid(lbPCName.Text, 1, 1))
                            words1(85) = ConvertDecToHex(Mid(lbPCName.Text, 2, 1))
                            words1(86) = ConvertDecToHex(Mid(lbPCName.Text, 3, 1))
                            words1(87) = ConvertDecToHex(Mid(lbPCName.Text, 4, 1))
                            words1(88) = ConvertDecToHex(Mid(lbPCName.Text, 5, 1))
                            words1(89) = ConvertDecToHex(Mid(lbPCName.Text, 6, 1))
                            words1(90) = ConvertDecToHex(Mid(lbPCName.Text, 7, 1))
                            words1(91) = ConvertDecToHex(Mid(lbPCName.Text, 8, 1))
                            words1(92) = ConvertDecToHex(Mid(lbPCName.Text, 9, 1))
                            words1(93) = ConvertDecToHex(Mid(lbPCName.Text, 10, 1))
                            words1(94) = ConvertDecToHex(Mid(lbPCName.Text, 11, 1))
                            words1(95) = ConvertDecToHex(Mid(lbPCName.Text, 12, 1))
                            words1(96) = ConvertDecToHex(Mid(lbPCName.Text, 13, 1))
                            words1(97) = ConvertDecToHex(Mid(lbPCName.Text, 14, 1))
                            words1(98) = ConvertDecToHex(Mid(lbPCName.Text, 15, 1))
                            words1(99) = ConvertDecToHex(Mid(lbPCName.Text, 16, 1))
                            words1(100) = ConvertDecToHex(Mid(lbPCName.Text, 17, 1))
                            words1(101) = ConvertDecToHex(Mid(lbPCName.Text, 18, 1))
                            words1(102) = ConvertDecToHex(Mid(lbPCName.Text, 19, 1))
                            words1(103) = ConvertDecToHex(Mid(lbPCName.Text, 20, 1))

                            words1(80) = Val(lbTankName.Text.Replace(",", ""))
                            words1(82) = Val(lbTankQty.Text.Replace(",", ""))
                            Exit Select
                        Case 2
                            words2(84) = ConvertDecToHex(Mid(lbPCName.Text, 1, 1))
                            words2(85) = ConvertDecToHex(Mid(lbPCName.Text, 2, 1))
                            words2(86) = ConvertDecToHex(Mid(lbPCName.Text, 3, 1))
                            words2(87) = ConvertDecToHex(Mid(lbPCName.Text, 4, 1))
                            words2(88) = ConvertDecToHex(Mid(lbPCName.Text, 5, 1))
                            words2(89) = ConvertDecToHex(Mid(lbPCName.Text, 6, 1))
                            words2(90) = ConvertDecToHex(Mid(lbPCName.Text, 7, 1))
                            words2(91) = ConvertDecToHex(Mid(lbPCName.Text, 8, 1))
                            words2(92) = ConvertDecToHex(Mid(lbPCName.Text, 9, 1))
                            words2(93) = ConvertDecToHex(Mid(lbPCName.Text, 10, 1))
                            words2(94) = ConvertDecToHex(Mid(lbPCName.Text, 11, 1))
                            words2(95) = ConvertDecToHex(Mid(lbPCName.Text, 12, 1))
                            words2(96) = ConvertDecToHex(Mid(lbPCName.Text, 13, 1))
                            words2(97) = ConvertDecToHex(Mid(lbPCName.Text, 14, 1))
                            words2(98) = ConvertDecToHex(Mid(lbPCName.Text, 15, 1))
                            words2(99) = ConvertDecToHex(Mid(lbPCName.Text, 16, 1))
                            words2(100) = ConvertDecToHex(Mid(lbPCName.Text, 17, 1))
                            words2(101) = ConvertDecToHex(Mid(lbPCName.Text, 18, 1))
                            words2(102) = ConvertDecToHex(Mid(lbPCName.Text, 19, 1))
                            words2(103) = ConvertDecToHex(Mid(lbPCName.Text, 20, 1))

                            words2(80) = Val(lbTankName.Text.Replace(",", ""))
                            words2(82) = Val(lbTankQty.Text.Replace(",", ""))
                            Exit Select
                        Case 3
                            words3(84) = ConvertDecToHex(Mid(lbPCName.Text, 1, 1))
                            words3(85) = ConvertDecToHex(Mid(lbPCName.Text, 2, 1))
                            words3(86) = ConvertDecToHex(Mid(lbPCName.Text, 3, 1))
                            words3(87) = ConvertDecToHex(Mid(lbPCName.Text, 4, 1))
                            words3(88) = ConvertDecToHex(Mid(lbPCName.Text, 5, 1))
                            words3(89) = ConvertDecToHex(Mid(lbPCName.Text, 6, 1))
                            words3(90) = ConvertDecToHex(Mid(lbPCName.Text, 7, 1))
                            words3(91) = ConvertDecToHex(Mid(lbPCName.Text, 8, 1))
                            words3(92) = ConvertDecToHex(Mid(lbPCName.Text, 9, 1))
                            words3(93) = ConvertDecToHex(Mid(lbPCName.Text, 10, 1))
                            words3(94) = ConvertDecToHex(Mid(lbPCName.Text, 11, 1))
                            words3(95) = ConvertDecToHex(Mid(lbPCName.Text, 12, 1))
                            words3(96) = ConvertDecToHex(Mid(lbPCName.Text, 13, 1))
                            words3(97) = ConvertDecToHex(Mid(lbPCName.Text, 14, 1))
                            words3(98) = ConvertDecToHex(Mid(lbPCName.Text, 15, 1))
                            words3(99) = ConvertDecToHex(Mid(lbPCName.Text, 16, 1))
                            words3(100) = ConvertDecToHex(Mid(lbPCName.Text, 17, 1))
                            words3(101) = ConvertDecToHex(Mid(lbPCName.Text, 18, 1))
                            words3(102) = ConvertDecToHex(Mid(lbPCName.Text, 19, 1))
                            words3(103) = ConvertDecToHex(Mid(lbPCName.Text, 20, 1))
                            'MsgBox(Val(lbTankName.Text.Replace(",", "")))
                            words3(80) = Val(lbTankName.Text.Replace(",", ""))
                            words3(82) = Val(lbTankQty.Text.Replace(",", ""))
                            'TextBox4.Text = words3(80) & " :: " & words2(82)
                            Exit Select
                        Case 4
                            words4(84) = ConvertDecToHex(Mid(lbPCName.Text, 1, 1))
                            words4(85) = ConvertDecToHex(Mid(lbPCName.Text, 2, 1))
                            words4(86) = ConvertDecToHex(Mid(lbPCName.Text, 3, 1))
                            words4(87) = ConvertDecToHex(Mid(lbPCName.Text, 4, 1))
                            words4(88) = ConvertDecToHex(Mid(lbPCName.Text, 5, 1))
                            words4(89) = ConvertDecToHex(Mid(lbPCName.Text, 6, 1))
                            words4(90) = ConvertDecToHex(Mid(lbPCName.Text, 7, 1))
                            words4(91) = ConvertDecToHex(Mid(lbPCName.Text, 8, 1))
                            words4(92) = ConvertDecToHex(Mid(lbPCName.Text, 9, 1))
                            words4(93) = ConvertDecToHex(Mid(lbPCName.Text, 10, 1))
                            words4(94) = ConvertDecToHex(Mid(lbPCName.Text, 11, 1))
                            words4(95) = ConvertDecToHex(Mid(lbPCName.Text, 12, 1))
                            words4(96) = ConvertDecToHex(Mid(lbPCName.Text, 13, 1))
                            words4(97) = ConvertDecToHex(Mid(lbPCName.Text, 14, 1))
                            words4(98) = ConvertDecToHex(Mid(lbPCName.Text, 15, 1))
                            words4(99) = ConvertDecToHex(Mid(lbPCName.Text, 16, 1))
                            words4(100) = ConvertDecToHex(Mid(lbPCName.Text, 17, 1))
                            words4(101) = ConvertDecToHex(Mid(lbPCName.Text, 18, 1))
                            words4(102) = ConvertDecToHex(Mid(lbPCName.Text, 19, 1))
                            words4(103) = ConvertDecToHex(Mid(lbPCName.Text, 20, 1))

                            words4(80) = Val(lbTankName.Text.Replace(",", ""))
                            words4(82) = Val(lbTankQty.Text.Replace(",", ""))
                            Exit Select
                        Case 5

                            'MsgBox(lbPCName.Text)
                            words5(84) = ConvertDecToHex(Mid(lbPCName.Text, 1, 1))
                            words5(85) = ConvertDecToHex(Mid(lbPCName.Text, 2, 1))
                            words5(86) = ConvertDecToHex(Mid(lbPCName.Text, 3, 1))
                            words5(87) = ConvertDecToHex(Mid(lbPCName.Text, 4, 1))
                            words5(88) = ConvertDecToHex(Mid(lbPCName.Text, 5, 1))
                            words5(89) = ConvertDecToHex(Mid(lbPCName.Text, 6, 1))
                            words5(90) = ConvertDecToHex(Mid(lbPCName.Text, 7, 1))
                            words5(91) = ConvertDecToHex(Mid(lbPCName.Text, 8, 1))
                            words5(92) = ConvertDecToHex(Mid(lbPCName.Text, 9, 1))
                            words5(93) = ConvertDecToHex(Mid(lbPCName.Text, 10, 1))
                            words5(94) = ConvertDecToHex(Mid(lbPCName.Text, 11, 1))
                            words5(95) = ConvertDecToHex(Mid(lbPCName.Text, 12, 1))
                            words5(96) = ConvertDecToHex(Mid(lbPCName.Text, 13, 1))
                            words5(97) = ConvertDecToHex(Mid(lbPCName.Text, 14, 1))
                            words5(98) = ConvertDecToHex(Mid(lbPCName.Text, 15, 1))
                            words5(99) = ConvertDecToHex(Mid(lbPCName.Text, 16, 1))
                            words5(100) = ConvertDecToHex(Mid(lbPCName.Text, 17, 1))
                            words5(101) = ConvertDecToHex(Mid(lbPCName.Text, 18, 1))
                            words5(102) = ConvertDecToHex(Mid(lbPCName.Text, 19, 1))
                            words5(103) = ConvertDecToHex(Mid(lbPCName.Text, 20, 1))

                            words5(80) = Val(lbTankName.Text.Replace(",", ""))
                            words5(82) = Val(lbTankQty.Text.Replace(",", ""))
                            Exit Select
                        Case 6
                            words6(84) = ConvertDecToHex(Mid(lbPCName.Text, 1, 1))
                            words6(85) = ConvertDecToHex(Mid(lbPCName.Text, 2, 1))
                            words6(86) = ConvertDecToHex(Mid(lbPCName.Text, 3, 1))
                            words6(87) = ConvertDecToHex(Mid(lbPCName.Text, 4, 1))
                            words6(88) = ConvertDecToHex(Mid(lbPCName.Text, 5, 1))
                            words6(89) = ConvertDecToHex(Mid(lbPCName.Text, 6, 1))
                            words6(90) = ConvertDecToHex(Mid(lbPCName.Text, 7, 1))
                            words6(91) = ConvertDecToHex(Mid(lbPCName.Text, 8, 1))
                            words6(92) = ConvertDecToHex(Mid(lbPCName.Text, 9, 1))
                            words6(93) = ConvertDecToHex(Mid(lbPCName.Text, 10, 1))
                            words6(94) = ConvertDecToHex(Mid(lbPCName.Text, 11, 1))
                            words6(95) = ConvertDecToHex(Mid(lbPCName.Text, 12, 1))
                            words6(96) = ConvertDecToHex(Mid(lbPCName.Text, 13, 1))
                            words6(97) = ConvertDecToHex(Mid(lbPCName.Text, 14, 1))
                            words6(98) = ConvertDecToHex(Mid(lbPCName.Text, 15, 1))
                            words6(99) = ConvertDecToHex(Mid(lbPCName.Text, 16, 1))
                            words6(100) = ConvertDecToHex(Mid(lbPCName.Text, 17, 1))
                            words6(101) = ConvertDecToHex(Mid(lbPCName.Text, 18, 1))
                            words6(102) = ConvertDecToHex(Mid(lbPCName.Text, 19, 1))
                            words6(103) = ConvertDecToHex(Mid(lbPCName.Text, 20, 1))

                            words6(80) = Val(lbTankName.Text.Replace(",", ""))
                            words6(82) = Val(lbTankQty.Text.Replace(",", ""))

                            Exit Select
                    End Select
                End With

                ' For i = 0 To subDS.Tables("master").Rows.Count - 1
                Dim strRMvalues As String = "0"
                With subDS.Tables("master").Rows(i)

                    rmName = .Item("BOM_Dtl_Name")
                    rmValve = subDS.Tables("master0").Rows(n).Item("Val_RM_Number")

                    rmCode = .Item("rm_Code")
                    rmValues = .Item("BOM_RM_Values")
                    If rmValve = "5" Or rmValve = "4" Then
                        strRMvalues = Format(rmValues, "#,##0.0")
                    Else
                        strRMvalues = Format(rmValues, "#,##0.0")
                    End If
                    rmUnit = .Item("BOM_Dtl_Unit_1")


                    dblTotal = (rmValues + dblTotal)
                End With

                anyData = New String() {n + 1, rmCode, rmName, strRMvalues, rmUnit, rmValve}
                lvi = New ListViewItem(anyData)
                'lsvBOM01.Items.Add(lvi)
                Dim txtstrRMvalues As New TextBox
                txtstrRMvalues.Text = strRMvalues
                'MsgBox("AAA")
                Select Case numScale
                    Case 1
                        lsvScale1.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words1(10) = 1
                            Else
                                words1(10) = 0
                            End If
                            words1(12) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words1(16) = rmCode
                            codeS1v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words1(18) = 1
                            Else
                                words1(18) = 0
                            End If
                            words1(20) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words1(24) = rmCode
                            codeS1V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words1(26) = 1
                            Else
                                words1(26) = 0
                            End If
                            words1(28) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words1(32) = rmCode
                            codeS1V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words1(34) = 1
                            Else
                                words1(34) = 0
                            End If
                            words1(36) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words1(40) = rmCode
                            codeS1V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words1(42) = 1
                            Else
                                words1(42) = 0
                            End If
                            words1(44) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words1(48) = rmCode
                            codeS1V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words1(50) = 1
                            Else
                                words1(50) = 0
                            End If
                            words1(52) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words1(56) = rmCode
                            codeS1V6 = rmCode
                        End If
                        Exit Select
                    Case 2
                        lsvScale2.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words2(10) = 1
                            Else
                                words2(10) = 0
                            End If
                            words2(12) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words2(16) = rmCode
                            codeS2v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words2(18) = 1
                            Else
                                words2(18) = 0
                            End If
                            words2(20) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words2(24) = rmCode
                            codeS2V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words2(26) = 1
                            Else
                                words2(26) = 0
                            End If
                            words2(28) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words2(32) = rmCode
                            codeS2V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words2(34) = 1
                            Else
                                words2(34) = 0
                            End If
                            words2(36) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words2(40) = rmCode
                            codeS2V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words2(42) = 1
                            Else
                                words2(42) = 0
                            End If
                            words2(44) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words2(48) = rmCode
                            codeS2V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words2(50) = 1
                            Else
                                words2(50) = 0
                            End If
                            words2(52) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words2(56) = rmCode
                            codeS2V6 = rmCode
                        End If
                        Exit Select
                    Case 3
                        lsvScale3.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words3(10) = 1
                            Else
                                words3(10) = 0
                            End If
                            words3(12) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words3(16) = rmCode
                            codeS3v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words3(18) = 1
                            Else
                                words3(18) = 0
                            End If
                            words3(20) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words3(24) = rmCode
                            codeS3V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words3(26) = 1
                            Else
                                words3(26) = 0
                            End If
                            words3(28) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words3(32) = rmCode
                            codeS3V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words3(34) = 1
                            Else
                                words3(34) = 0
                            End If
                            words3(36) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words3(40) = rmCode
                            codeS3V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words3(42) = 1
                            Else
                                words3(42) = 0
                            End If
                            words3(44) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words3(48) = rmCode
                            codeS3V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words3(50) = 1
                            Else
                                words3(50) = 0
                            End If
                            words3(52) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words3(56) = rmCode
                            codeS3V6 = rmCode
                        End If
                        Exit Select
                    Case 4
                        lsvScale4.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words4(10) = 1
                            Else
                                words4(10) = 0
                            End If
                            words4(12) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words4(16) = rmCode
                            codeS4v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words4(18) = 1
                            Else
                                words4(18) = 0
                            End If
                            words4(20) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words4(24) = rmCode
                            codeS4V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words4(26) = 1
                            Else
                                words4(26) = 0
                            End If
                            words4(28) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words4(32) = rmCode
                            codeS4V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words4(34) = 1
                            Else
                                words4(34) = 0
                            End If
                            words4(36) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words4(40) = rmCode
                            codeS4V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words4(42) = 1
                            Else
                                words4(42) = 0
                            End If
                            words4(44) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words4(48) = rmCode
                            codeS4V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words4(50) = 1
                            Else
                                words4(50) = 0
                            End If
                            words4(52) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words4(56) = rmCode
                            codeS4V6 = rmCode
                        End If
                        Exit Select
                    Case 5
                        lsvScale5.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words5(10) = 1
                            Else
                                words5(10) = 0
                            End If
                            words5(12) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words5(16) = rmCode
                            codeS5v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words5(18) = 1
                            Else
                                words5(18) = 0
                            End If
                            words5(20) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words5(24) = rmCode
                            codeS5V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words5(26) = 1
                            Else
                                words5(26) = 0
                            End If
                            words5(28) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words5(32) = rmCode
                            codeS5V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words5(34) = 1
                            Else
                                words5(34) = 0
                            End If
                            words5(36) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words5(40) = rmCode
                            codeS5V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words5(42) = 1
                            Else
                                words5(42) = 0
                            End If
                            words5(44) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words5(48) = rmCode
                            codeS5V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words5(50) = 1
                            Else
                                words5(50) = 0
                            End If
                            words5(52) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words5(56) = rmCode
                            codeS5V6 = rmCode
                        End If
                        Exit Select
                    Case 6
                        lsvScale6.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words6(10) = 1
                            Else
                                words6(10) = 0
                            End If
                            words6(12) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words6(16) = rmCode
                            codeS6v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words6(18) = 1
                            Else
                                words6(18) = 0
                            End If
                            words6(20) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words6(24) = rmCode
                            codeS6V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words6(26) = 1
                            Else
                                words6(26) = 0
                            End If
                            words6(28) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words6(32) = rmCode
                            codeS6V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words6(34) = 1
                            Else
                                words6(34) = 0
                            End If
                            words6(36) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words6(40) = rmCode
                            codeS6V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words6(42) = 1
                            Else
                                words6(42) = 0
                            End If
                            words6(44) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words6(48) = rmCode
                            codeS6V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words6(50) = 1
                            Else
                                words6(50) = 0
                            End If
                            words6(52) = Val(txtstrRMvalues.Text.Replace(".", ""))
                            words6(56) = rmCode
                            codeS6V6 = rmCode
                        End If
                        Exit Select
                End Select

                If chkRow = 0 Then
                    lvi.BackColor = Color.White
                    lvi.ForeColor = Color.Black
                    chkRow = 1

                ElseIf chkRow = 1 Then
                    lvi.BackColor = Color.PaleGoldenrod
                    lvi.ForeColor = Color.Black
                    chkRow = 0

                End If
            Else
                rmName = subDS.Tables("master0").Rows(n).Item("BOM_Dtl_Name")
                rmCode = subDS.Tables("master0").Rows(n).Item("rm_Code")
                rmValues = "0" '.Item("BOM_RM_Values")
                rmUnit = "-" '.Item("BOM_Dtl_Unit_1")
                rmValve = subDS.Tables("master0").Rows(n).Item("Val_RM_Number")
                anyData = New String() {n + 1, rmCode, rmName, rmValues, rmUnit, rmValve}

                lvi = New ListViewItem(anyData)
                'lsvBOM01.Items.Add(lvi)
                Dim txtrmValues As New TextBox
                txtrmValues.Text = rmValues
                'MsgBox("BBB")
                Select Case numScale
                    Case 1
                        lsvScale1.Items.Add(lvi)
                        'MsgBox("SSS")
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words1(10) = 1
                            Else
                                words1(10) = 0
                            End If
                            words1(12) = Val(txtrmValues.Text.Replace(".", ""))
                            words1(16) = rmCode
                            codeS1v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words1(18) = 1
                            Else
                                words1(18) = 0
                            End If
                            words1(20) = Val(txtrmValues.Text.Replace(".", ""))
                            words1(24) = rmCode
                            codeS1V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words1(26) = 1
                            Else
                                words1(26) = 0
                            End If
                            words1(28) = Val(txtrmValues.Text.Replace(".", ""))
                            words1(32) = rmCode
                            codeS1V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words1(34) = 1
                            Else
                                words1(34) = 0
                            End If
                            words1(36) = Val(txtrmValues.Text.Replace(".", ""))
                            words1(40) = rmCode
                            codeS1V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words1(42) = 1
                            Else
                                words1(42) = 0
                            End If
                            words1(44) = Val(txtrmValues.Text.Replace(".", ""))
                            words1(48) = rmCode
                            codeS1V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words1(50) = 1
                            Else
                                words1(50) = 0
                            End If
                            words1(52) = Val(txtrmValues.Text.Replace(".", ""))
                            words1(56) = rmCode
                            codeS1V6 = rmCode
                        End If
                        Exit Select
                    Case 2
                        lsvScale2.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words2(10) = 1
                            Else
                                words2(10) = 0
                            End If
                            words2(12) = Val(txtrmValues.Text.Replace(".", ""))
                            words2(16) = rmCode
                            codeS2v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words2(18) = 1
                            Else
                                words2(18) = 0
                            End If
                            words2(20) = Val(txtrmValues.Text.Replace(".", ""))
                            words2(24) = rmCode
                            codeS2V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words2(26) = 1
                            Else
                                words2(26) = 0
                            End If
                            words2(28) = Val(txtrmValues.Text.Replace(".", ""))
                            words2(32) = rmCode
                            codeS2V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words2(34) = 1
                            Else
                                words2(34) = 0
                            End If
                            words2(36) = Val(txtrmValues.Text.Replace(".", ""))
                            words2(40) = rmCode
                            codeS2V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words2(42) = 1
                            Else
                                words2(42) = 0
                            End If
                            words2(44) = Val(txtrmValues.Text.Replace(".", ""))
                            words2(48) = rmCode
                            codeS2V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words2(50) = 1
                            Else
                                words2(50) = 0
                            End If
                            words2(52) = Val(txtrmValues.Text.Replace(".", ""))
                            words2(56) = rmCode
                            codeS2V6 = rmCode
                        End If
                        Exit Select
                    Case 3
                        lsvScale3.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words3(10) = 1
                            Else
                                words3(10) = 0
                            End If
                            words3(12) = Val(txtrmValues.Text.Replace(".", ""))
                            words3(16) = rmCode
                            codeS3v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words3(18) = 1
                            Else
                                words3(18) = 0
                            End If
                            words3(20) = Val(txtrmValues.Text.Replace(".", ""))
                            words3(24) = rmCode
                            codeS3V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words3(26) = 1
                            Else
                                words3(26) = 0
                            End If
                            words3(28) = Val(txtrmValues.Text.Replace(".", ""))
                            words3(32) = rmCode
                            codeS3V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words3(34) = 1
                            Else
                                words3(34) = 0
                            End If
                            words3(36) = Val(txtrmValues.Text.Replace(".", ""))
                            words3(40) = rmCode
                            codeS3V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words3(42) = 1
                            Else
                                words3(42) = 0
                            End If
                            words3(44) = Val(txtrmValues.Text.Replace(".", ""))
                            words3(48) = rmCode
                            codeS3V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words3(50) = 1
                            Else
                                words3(50) = 0
                            End If
                            words3(52) = Val(txtrmValues.Text.Replace(".", ""))
                            words3(56) = rmCode
                            codeS3V6 = rmCode
                        End If
                        Exit Select
                    Case 4
                        lsvScale4.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words4(10) = 1
                            Else
                                words4(10) = 0
                            End If
                            words4(12) = Val(txtrmValues.Text.Replace(".", ""))
                            words4(16) = rmCode
                            codeS4v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words4(18) = 1
                            Else
                                words4(18) = 0
                            End If
                            words4(20) = Val(txtrmValues.Text.Replace(".", ""))
                            words4(24) = rmCode
                            codeS4V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words4(26) = 1
                            Else
                                words4(26) = 0
                            End If
                            words4(28) = Val(txtrmValues.Text.Replace(".", ""))
                            words4(32) = rmCode
                            codeS4V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words4(34) = 1
                            Else
                                words4(34) = 0
                            End If
                            words4(36) = Val(txtrmValues.Text.Replace(".", ""))
                            words4(40) = rmCode
                            codeS4V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words4(42) = 1
                            Else
                                words4(42) = 0
                            End If
                            words4(44) = Val(txtrmValues.Text.Replace(".", ""))
                            words4(48) = rmCode
                            codeS4V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words4(50) = 1
                            Else
                                words4(50) = 0
                            End If
                            words4(52) = Val(txtrmValues.Text.Replace(".", ""))
                            words4(56) = rmCode
                            codeS4V6 = rmCode
                        End If
                        Exit Select
                    Case 5
                        lsvScale5.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words5(10) = 1
                            Else
                                words5(10) = 0
                            End If
                            words5(12) = Val(txtrmValues.Text.Replace(".", ""))
                            words5(16) = rmCode
                            codeS5v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words5(18) = 1
                            Else
                                words5(18) = 0
                            End If
                            words5(20) = Val(txtrmValues.Text.Replace(".", ""))
                            words5(24) = rmCode
                            codeS5V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words5(26) = 1
                            Else
                                words5(26) = 0
                            End If
                            words5(28) = Val(txtrmValues.Text.Replace(".", ""))
                            words5(32) = rmCode
                            codeS5V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words5(34) = 1
                            Else
                                words5(34) = 0
                            End If
                            words5(36) = Val(txtrmValues.Text.Replace(".", ""))
                            words5(40) = rmCode
                            codeS5V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words5(42) = 1
                            Else
                                words5(42) = 0
                            End If
                            words5(44) = Val(txtrmValues.Text.Replace(".", ""))
                            words5(48) = rmCode
                            codeS5V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words5(50) = 1
                            Else
                                words5(50) = 0
                            End If
                            words5(52) = Val(txtrmValues.Text.Replace(".", ""))
                            words5(56) = rmCode
                            codeS5V6 = rmCode
                        End If
                        Exit Select
                    Case 6
                        lsvScale6.Items.Add(lvi)
                        If rmValve = 1 Then
                            If rmValues > 0 Then
                                words6(10) = 1
                            Else
                                words6(10) = 0
                            End If
                            words6(12) = Val(txtrmValues.Text.Replace(".", ""))
                            words6(16) = rmCode
                            codeS6v1 = rmCode
                        End If
                        If rmValve = 2 Then
                            If rmValues > 0 Then
                                words6(18) = 1
                            Else
                                words6(18) = 0
                            End If
                            words6(20) = Val(txtrmValues.Text.Replace(".", ""))
                            words6(24) = rmCode
                            codeS6V2 = rmCode
                        End If
                        If rmValve = 3 Then
                            If rmValues > 0 Then
                                words6(26) = 1
                            Else
                                words6(26) = 0
                            End If
                            words6(28) = Val(txtrmValues.Text.Replace(".", ""))
                            words6(32) = rmCode
                            codeS6V3 = rmCode
                        End If
                        If rmValve = 4 Then
                            If rmValues > 0 Then
                                words6(34) = 1
                            Else
                                words6(34) = 0
                            End If
                            words6(36) = Val(txtrmValues.Text.Replace(".", ""))
                            words6(40) = rmCode
                            codeS6V4 = rmCode
                        End If
                        If rmValve = 5 Then
                            If rmValues > 0 Then
                                words6(42) = 1
                            Else
                                words6(42) = 0
                            End If
                            words6(44) = Val(txtrmValues.Text.Replace(".", ""))
                            words6(48) = rmCode
                            codeS6V5 = rmCode
                        End If
                        If rmValve = 6 Then
                            If rmValues > 0 Then
                                words6(50) = 1
                            Else
                                words6(50) = 0
                            End If
                            words6(52) = Val(txtrmValues.Text.Replace(".", ""))
                            words6(56) = rmCode
                            codeS6V6 = rmCode
                        End If
                        Exit Select
                End Select

                If chkRow = 0 Then
                    lvi.BackColor = Color.White
                    lvi.ForeColor = Color.Black
                    chkRow = 1

                ElseIf chkRow = 1 Then
                    lvi.BackColor = Color.PaleGoldenrod
                    lvi.ForeColor = Color.Black
                    chkRow = 0

                End If
            End If
            nrec = n + 1
        Next

        If numScale = 1 Then
            lbTotalChkScale1.Text = dblTotal.ToString("#,##0.0")
            txtWeightRunScale1.Text = lbTotalChkScale1.Text
            nrec1 = nrec
        End If
        If numScale = 2 Then
            lbTotalChkScale2.Text = dblTotal.ToString("#,##0.0")
            txtWeightRunScale2.Text = lbTotalChkScale2.Text
            nrec2 = nrec
        End If
        If numScale = 3 Then
            lbTotalChkScale3.Text = dblTotal.ToString("#,##0.0")
            txtWeightRunScale3.Text = lbTotalChkScale3.Text
            nrec3 = nrec
        End If
        If numScale = 4 Then
            lbTotalChkScale4.Text = dblTotal.ToString("#,##0.0")
            txtWeightRunScale4.Text = lbTotalChkScale4.Text
            nrec4 = nrec
        End If
        If numScale = 5 Then
            lbTotalChkScale5.Text = dblTotal.ToString("#,##0.0")
            txtWeightRunScale5.Text = lbTotalChkScale5.Text
            nrec5 = nrec
        End If
        If numScale = 6 Then
            lbTotalChkScale6.Text = dblTotal.ToString("#,##0.0")
            txtWeightRunScale6.Text = lbTotalChkScale6.Text
            nrec6 = nrec
        End If
    End Sub

    Sub runFindScale(intScale As Integer)
        Dim strDocNO As String = ""
        Select Case intScale
            Case 1
                strDocNO = Trim(txtBarScale1.Text)
                Exit Select
            Case 2
                strDocNO = Trim(txtBarScale2.Text)
                Exit Select
            Case 3
                strDocNO = Trim(txtBarScale3.Text)
                Exit Select
            Case 4
                strDocNO = Trim(txtBarScale4.Text)
                Exit Select
            Case 5
                strDocNO = Trim(txtBarScale5.Text)
                Exit Select
            Case 6
                strDocNO = Trim(txtBarScale6.Text)
                Exit Select
        End Select
        'strDocNO = Trim(txtBarScale1.Text)
        showData(strDocNO, intScale)
    End Sub
    ' ------------------------------------------------------------------------
    ' Generate new number of text boxes
    ' ------------------------------------------------------------------------
    Private Sub ResizeData1()
        ' Create as many textboxes as fit into window
        grpData1.Controls.Clear()
        Dim x As Integer = 0
        Dim y As Integer = 10
        Dim z As Integer = 20
        Dim w As Integer = 1
        'While y < grpData.Size.Width - 100
        While y < grpData1.Size.Width
            labData1 = New Label()
            grpData1.Controls.Add(labData1)
            labData1.Size = New System.Drawing.Size(30, 20)
            labData1.Location = New System.Drawing.Point(y, z)
            labData1.Text = Convert.ToString(x + 1)

            txtdata1 = New TextBox()
            grpData1.Controls.Add(txtdata1)
            txtdata1.Size = New System.Drawing.Size(50, 20)
            txtdata1.Location = New System.Drawing.Point(y + 30, z)
            txtdata1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            txtdata1.Tag = (x + 1)

            x += 1
            z = z + txtdata1.Size.Height + 5
            If z > grpData1.Size.Height - 40 Then
                y = y + 100
                z = 20
            End If
        End While
    End Sub

    Private Sub ResizeData2()
        ' Create as many textboxes as fit into window
        grpData2.Controls.Clear()
        Dim x As Integer = 0
        Dim y As Integer = 10
        Dim z As Integer = 20
        Dim w As Integer = 1
        'While y < grpData.Size.Width - 100
        While y < grpData2.Size.Width
            labData2 = New Label()
            grpData2.Controls.Add(labData2)
            labData2.Size = New System.Drawing.Size(30, 20)
            labData2.Location = New System.Drawing.Point(y, z)
            labData2.Text = Convert.ToString(x + 1)

            txtdata2 = New TextBox()
            grpData2.Controls.Add(txtdata2)
            txtdata2.Size = New System.Drawing.Size(50, 20)
            txtdata2.Location = New System.Drawing.Point(y + 30, z)
            txtdata2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            txtdata2.Tag = (x + 1)

            x += 1
            z = z + txtdata2.Size.Height + 5
            If z > grpData2.Size.Height - 40 Then
                y = y + 100
                z = 20
            End If
        End While
    End Sub

    Private Sub ResizeData3()
        ' Create as many textboxes as fit into window
        grpData3.Controls.Clear()
        Dim x As Integer = 0
        Dim y As Integer = 10
        Dim z As Integer = 20
        Dim w As Integer = 1
        'While y < grpData.Size.Width - 100
        While y < grpData3.Size.Width
            labData3 = New Label()
            grpData3.Controls.Add(labData3)
            labData3.Size = New System.Drawing.Size(30, 20)
            labData3.Location = New System.Drawing.Point(y, z)
            labData3.Text = Convert.ToString(x + 1)

            txtdata3 = New TextBox()
            grpData3.Controls.Add(txtdata3)
            txtdata3.Size = New System.Drawing.Size(50, 20)
            txtdata3.Location = New System.Drawing.Point(y + 30, z)
            txtdata3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            txtdata3.Tag = (x + 1)

            x += 1
            z = z + txtdata3.Size.Height + 5
            If z > grpData3.Size.Height - 40 Then
                y = y + 100
                z = 20
            End If
        End While
    End Sub

    Private Sub ResizeData4()
        ' Create as many textboxes as fit into window
        grpData4.Controls.Clear()
        Dim x As Integer = 0
        Dim y As Integer = 10
        Dim z As Integer = 20
        Dim w As Integer = 1
        'While y < grpData.Size.Width - 100
        While y < grpData4.Size.Width
            labData4 = New Label()
            grpData4.Controls.Add(labData4)
            labData4.Size = New System.Drawing.Size(30, 20)
            labData4.Location = New System.Drawing.Point(y, z)
            labData4.Text = Convert.ToString(x + 1)

            txtdata4 = New TextBox()
            grpData4.Controls.Add(txtdata4)
            txtdata4.Size = New System.Drawing.Size(50, 20)
            txtdata4.Location = New System.Drawing.Point(y + 30, z)
            txtdata4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            txtdata4.Tag = (x + 1)

            x += 1
            z = z + txtdata4.Size.Height + 5
            If z > grpData4.Size.Height - 40 Then
                y = y + 100
                z = 20
            End If
        End While
    End Sub

    Private Sub ResizeData5()
        ' Create as many textboxes as fit into window
        grpData5.Controls.Clear()
        Dim x As Integer = 0
        Dim y As Integer = 10
        Dim z As Integer = 20
        Dim w As Integer = 1
        'While y < grpData.Size.Width - 100
        While y < grpData5.Size.Width
            labData5 = New Label()
            grpData5.Controls.Add(labData5)
            labData5.Size = New System.Drawing.Size(30, 20)
            labData5.Location = New System.Drawing.Point(y, z)
            labData5.Text = Convert.ToString(x + 1)

            txtdata5 = New TextBox()
            grpData5.Controls.Add(txtdata5)
            txtdata5.Size = New System.Drawing.Size(50, 20)
            txtdata5.Location = New System.Drawing.Point(y + 30, z)
            txtdata5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            txtdata5.Tag = (x + 1)

            x += 1
            z = z + txtdata5.Size.Height + 5
            If z > grpData5.Size.Height - 40 Then
                y = y + 100
                z = 20
            End If
        End While
    End Sub

    Private Sub ResizeData6()
        ' Create as many textboxes as fit into window
        grpData6.Controls.Clear()
        Dim x As Integer = 0
        Dim y As Integer = 10
        Dim z As Integer = 20
        Dim w As Integer = 1
        'While y < grpData.Size.Width - 100
        While y < grpData6.Size.Width
            labData6 = New Label()
            grpData6.Controls.Add(labData6)
            labData6.Size = New System.Drawing.Size(30, 20)
            labData6.Location = New System.Drawing.Point(y, z)
            labData6.Text = Convert.ToString(x + 1)

            txtdata6 = New TextBox()
            grpData6.Controls.Add(txtdata6)
            txtdata6.Size = New System.Drawing.Size(50, 20)
            txtdata6.Location = New System.Drawing.Point(y + 30, z)
            txtdata6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            txtdata6.Tag = (x + 1)

            x += 1
            z = z + txtdata6.Size.Height + 5
            If z > grpData6.Size.Height - 40 Then
                y = y + 100
                z = 20
            End If
        End While
    End Sub

    Private Sub FillNullwords(ByVal ncscale As Integer)
        If ncscale = 0 Then
            words1 = New Integer(Val(txtSize.Text)) {}
            words2 = New Integer(Val(txtSize.Text)) {}
            words3 = New Integer(Val(txtSize.Text)) {}
            words4 = New Integer(Val(txtSize.Text)) {}
            words5 = New Integer(Val(txtSize.Text)) {}
            words6 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words1(v) = 0
                words2(v) = 0
                words3(v) = 0
                words4(v) = 0
                words5(v) = 0
                words6(v) = 0
            Next
            'MsgBox(words1(10))
        End If

        If ncscale = 1 Then
            words1 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words1(v) = 0
            Next
        End If

        If ncscale = 2 Then
            words2 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words2(v) = 0
                'MsgBox(v)
            Next
        End If

        If ncscale = 3 Then
            words3 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words3(v) = 0
            Next
        End If

        If ncscale = 4 Then
            words4 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words4(v) = 0
            Next
        End If

        If ncscale = 5 Then
            words5 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words5(v) = 0
            Next
        End If

        If ncscale = 6 Then
            words6 = New Integer(Val(txtSize.Text)) {}
            For v As Integer = 0 To Val(txtSize.Text)
                words6(v) = 0
            Next
        End If
    End Sub

    Private Sub ShowDataToGridView()
        Dim DS As New DataSet
        Dim DA As SqlClient.SqlDataAdapter

        Dim lsStr As String = ""
        lsStr = "Select TOP 100 * From BOMmastF Order By BOM_RM_Updatetime DESC"

        lsvSummaryScale.Rows.Clear()
        DA = New SqlClient.SqlDataAdapter(lsStr, Conn)
        DA.Fill(DS, "Emp")

        Dim n As Integer
        For n = 0 To DS.Tables("Emp").Rows.Count - 1
            With DS.Tables("Emp").Rows(n)
                lsvSummaryScale.Rows.Add()
                lsvSummaryScale.Rows(n).Cells(0).Value = .Item(0)
                lsvSummaryScale.Rows(n).Cells(1).Value = .Item(1)
                lsvSummaryScale.Rows(n).Cells(2).Value = .Item(2)
                lsvSummaryScale.Rows(n).Cells(3).Value = Mid(.Item(6).ToString, 1, 20)
                lsvSummaryScale.Rows(n).Cells(4).Value = .Item(4)
                lsvSummaryScale.Rows(n).Cells(5).Value = .Item(5)
            End With
        Next
    End Sub

    Function checkSaveDouble(strCode As String) As Integer
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter

        subDA = New SqlClient.SqlDataAdapter(strCode, Conn)
        subDA.Fill(subDS, "master")

        Return subDS.Tables("master").Rows.Count - 1
    End Function

    Private Sub SaveDataScale(ByVal nsScale As Integer)
        Dim csql1, csql2, csql3, csql4, csql5 As String
        Select Case nsScale
            Case 1
                Try
                    'Scale#1
                    'Valve1
                    csql1 = ""
                    csql1 = csql1 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale1.Text & "') and (BOM_RM_Code = '" & codeS1v1 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub1 & ") and (BOM_RM_Scales = 1) "
                    If checkSaveDouble(csql1) < 0 Then
                        If txtBarScale1.Text <> "" And codeS1v1 <> "" Then
                            sqlStr1 = ""
                            sqlStr1 = sqlStr1 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr1 = sqlStr1 & " VALUES('" & txtBarScale1.Text & "', '" & codeS1v1 & "', " & wordx1(14) & ", getdate(), " & coub1 & ", 1, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr1, "")
                        End If
                    End If

                    'Valve2
                    csql2 = ""
                    csql2 = csql2 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale1.Text & "') and (BOM_RM_Code = '" & codeS1V2 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub2 & ") and (BOM_RM_Scales = 1) "
                    If checkSaveDouble(csql2) < 0 Then
                        If txtBarScale1.Text <> "" And codeS1V2 <> "" Then
                            sqlStr1 = ""
                            sqlStr1 = sqlStr1 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr1 = sqlStr1 & " VALUES('" & txtBarScale1.Text & "', '" & codeS1V2 & "', " & wordx1(22) & ", getdate(), " & coub1 & ", 1, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr1, "")
                        End If
                    End If

                    'Valve3
                    csql3 = ""
                    csql3 = csql3 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale1.Text & "') and (BOM_RM_Code = '" & codeS1V3 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub3 & ") and (BOM_RM_Scales = 1) "
                    If checkSaveDouble(csql3) < 0 Then
                        If txtBarScale1.Text <> "" And codeS1V3 <> "" Then
                            sqlStr1 = ""
                            sqlStr1 = sqlStr1 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr1 = sqlStr1 & " VALUES('" & txtBarScale1.Text & "', '" & codeS1V3 & "', " & wordx1(30) & ", getdate(), " & coub1 & ", 1, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr1, "")
                        End If
                    End If

                    'Valve4
                    csql4 = ""
                    csql4 = csql4 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale1.Text & "') and (BOM_RM_Code = '" & codeS1V4 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub4 & ") and (BOM_RM_Scales = 1) "
                    If checkSaveDouble(csql4) < 0 Then
                        If txtBarScale1.Text <> "" And codeS1V4 <> "" Then
                            sqlStr1 = ""
                            sqlStr1 = sqlStr1 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr1 = sqlStr1 & " VALUES('" & txtBarScale1.Text & "', '" & codeS1V4 & "', " & wordx1(38) & ", getdate(), " & coub1 & ", 1, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr1, "")
                        End If
                    End If

                    'Valve5
                    csql5 = ""
                    csql5 = csql5 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale1.Text & "') and (BOM_RM_Code = '" & codeS1V5 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub5 & ") and (BOM_RM_Scales = 1) "
                    If checkSaveDouble(csql5) < 0 Then
                        If txtBarScale1.Text <> "" And codeS1V5 <> "" Then
                            sqlStr1 = ""
                            sqlStr1 = sqlStr1 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr1 = sqlStr1 & " VALUES('" & txtBarScale1.Text & "', '" & codeS1V5 & "', " & wordx1(46) & ", getdate(), " & coub1 & ", 1, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr1, "")
                        End If
                    End If

                    ShowDataToGridView()
                Catch ex As Exception

                End Try
                Exit Select
            Case 2
                Try
                    'Scale#2
                    'Valve1
                    csql1 = ""
                    csql1 = csql1 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale2.Text & "') and (BOM_RM_Code = '" & codeS2v1 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub1 & ") and (BOM_RM_Scales = 2) "
                    If checkSaveDouble(csql1) < 0 Then
                        If txtBarScale2.Text <> "" And codeS2v1 <> "" Then
                            sqlStr2 = ""
                            sqlStr2 = sqlStr2 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr2 = sqlStr2 & " VALUES('" & txtBarScale2.Text & "', '" & codeS2v1 & "', " & wordx2(14) & ", getdate(), " & coub2 & ", 2, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr2, "")
                        End If
                    End If

                    'Valve2
                    csql2 = ""
                    csql2 = csql2 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale2.Text & "') and (BOM_RM_Code = '" & codeS2V2 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub2 & ") and (BOM_RM_Scales = 2) "
                    If checkSaveDouble(csql2) < 0 Then
                        If txtBarScale2.Text <> "" And codeS2V2 <> "" Then
                            sqlStr2 = ""
                            sqlStr2 = sqlStr2 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr2 = sqlStr2 & " VALUES('" & txtBarScale2.Text & "', '" & codeS2V2 & "', " & wordx2(22) & ", getdate(), " & coub2 & ", 2, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr2, "")
                        End If
                    End If

                    'Valve3
                    csql3 = ""
                    csql3 = csql3 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale2.Text & "') and (BOM_RM_Code = '" & codeS2V3 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub3 & ") and (BOM_RM_Scales = 2) "
                    If checkSaveDouble(csql3) < 0 Then
                        If txtBarScale2.Text <> "" And codeS2V3 <> "" Then
                            sqlStr2 = ""
                            sqlStr2 = sqlStr2 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr2 = sqlStr2 & " VALUES('" & txtBarScale2.Text & "', '" & codeS2V3 & "', " & wordx2(30) & ", getdate(), " & coub2 & ", 2, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr2, "")
                        End If
                    End If

                    'Valve4
                    csql4 = ""
                    csql4 = csql4 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale2.Text & "') and (BOM_RM_Code = '" & codeS2V4 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub4 & ") and (BOM_RM_Scales = 2) "
                    If checkSaveDouble(csql4) < 0 Then
                        If txtBarScale2.Text <> "" And codeS2V4 <> "" Then
                            sqlStr2 = ""
                            sqlStr2 = sqlStr2 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr2 = sqlStr2 & " VALUES('" & txtBarScale2.Text & "', '" & codeS2V4 & "', " & wordx2(38) & ", getdate(), " & coub2 & ", 2, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr2, "")
                        End If
                    End If

                    'Valve5
                    csql5 = ""
                    csql5 = csql5 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale2.Text & "') and (BOM_RM_Code = '" & codeS2V5 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub5 & ") and (BOM_RM_Scales = 2) "
                    If checkSaveDouble(csql5) < 0 Then
                        If txtBarScale2.Text <> "" And codeS2V5 <> "" Then
                            sqlStr2 = ""
                            sqlStr2 = sqlStr2 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr2 = sqlStr2 & " VALUES('" & txtBarScale2.Text & "', '" & codeS2V5 & "', " & wordx2(46) & ", getdate(), " & coub2 & ", 2, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr2, "")
                        End If
                    End If

                    ShowDataToGridView()
                Catch ex As Exception

                End Try
                Exit Select
            Case 3
                Try
                    'Scale#3
                    'Valve1
                    csql1 = ""
                    csql1 = csql1 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale3.Text & "') and (BOM_RM_Code = '" & codeS3v1 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub1 & ") and (BOM_RM_Scales = 3) "
                    If checkSaveDouble(csql1) < 0 Then
                        If txtBarScale3.Text <> "" And codeS3v1 <> "" Then
                            sqlStr3 = ""
                            sqlStr3 = sqlStr3 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr3 = sqlStr3 & " VALUES('" & txtBarScale3.Text & "', '" & codeS3v1 & "', " & wordx3(14) & ", getdate(), " & coub3 & ", 3, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr3, "")
                        End If
                    End If

                    'Valve2
                    csql2 = ""
                    csql2 = csql2 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale3.Text & "') and (BOM_RM_Code = '" & codeS3V2 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub2 & ") and (BOM_RM_Scales = 3) "
                    If checkSaveDouble(csql2) < 0 Then
                        If txtBarScale3.Text <> "" And codeS3V2 <> "" Then
                            sqlStr3 = ""
                            sqlStr3 = sqlStr3 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr3 = sqlStr3 & " VALUES('" & txtBarScale3.Text & "', '" & codeS3V2 & "', " & wordx3(22) & ", getdate(), " & coub3 & ", 3, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr3, "")
                        End If
                    End If

                    'Valve3
                    csql3 = ""
                    csql3 = csql3 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale3.Text & "') and (BOM_RM_Code = '" & codeS3V3 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub3 & ") and (BOM_RM_Scales = 3) "
                    If checkSaveDouble(csql3) < 0 Then
                        If txtBarScale3.Text <> "" And codeS3V3 <> "" Then
                            sqlStr3 = ""
                            sqlStr3 = sqlStr3 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr3 = sqlStr3 & " VALUES('" & txtBarScale3.Text & "', '" & codeS3V3 & "', " & wordx3(30) & ", getdate(), " & coub3 & ", 3, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr3, "")
                        End If
                    End If

                    'Valve4
                    csql4 = ""
                    csql4 = csql4 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale3.Text & "') and (BOM_RM_Code = '" & codeS3V4 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub4 & ") and (BOM_RM_Scales = 3) "
                    If checkSaveDouble(csql4) < 0 Then
                        If txtBarScale3.Text <> "" And codeS3V4 <> "" Then
                            sqlStr3 = ""
                            sqlStr3 = sqlStr3 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr3 = sqlStr3 & " VALUES('" & txtBarScale3.Text & "', '" & codeS3V4 & "', " & wordx3(38) & ", getdate(), " & coub3 & ", 3, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr3, "")
                        End If
                    End If

                    'Valve5
                    csql5 = ""
                    csql5 = csql5 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale3.Text & "') and (BOM_RM_Code = '" & codeS3V5 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub5 & ") and (BOM_RM_Scales = 3) "
                    If checkSaveDouble(csql5) < 0 Then
                        If txtBarScale3.Text <> "" And codeS3V5 <> "" Then
                            sqlStr3 = ""
                            sqlStr3 = sqlStr3 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr3 = sqlStr3 & " VALUES('" & txtBarScale3.Text & "', '" & codeS3V5 & "', " & wordx3(46) & ", getdate(), " & coub3 & ", 3, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr3, "")
                        End If
                    End If

                    ShowDataToGridView()
                Catch ex As Exception

                End Try
                Exit Select
            Case 4
                Try
                    'Scale#4
                    'Valve1
                    csql1 = ""
                    csql1 = csql1 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale4.Text & "') and (BOM_RM_Code = '" & codeS4v1 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub1 & ") and (BOM_RM_Scales = 4) "
                    If checkSaveDouble(csql1) < 0 Then
                        If txtBarScale4.Text <> "" And codeS4v1 <> "" Then
                            sqlStr4 = ""
                            sqlStr4 = sqlStr4 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr4 = sqlStr4 & " VALUES('" & txtBarScale4.Text & "', '" & codeS4v1 & "', " & wordx4(14) & ", getdate(), " & coub4 & ", 4, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr4, "")
                        End If
                    End If

                    'Valve2
                    csql2 = ""
                    csql2 = csql2 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale4.Text & "') and (BOM_RM_Code = '" & codeS4V2 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub2 & ") and (BOM_RM_Scales = 4) "
                    If checkSaveDouble(csql2) < 0 Then
                        If txtBarScale4.Text <> "" And codeS4V2 <> "" Then
                            sqlStr4 = ""
                            sqlStr4 = sqlStr4 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr4 = sqlStr4 & " VALUES('" & txtBarScale4.Text & "', '" & codeS4V2 & "', " & wordx4(22) & ", getdate(), " & coub4 & ", 4, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr4, "")
                        End If
                    End If

                    'Valve3
                    csql3 = ""
                    csql3 = csql3 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale4.Text & "') and (BOM_RM_Code = '" & codeS4V3 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub3 & ") and (BOM_RM_Scales = 4) "
                    If checkSaveDouble(csql3) < 0 Then
                        If txtBarScale4.Text <> "" And codeS4V3 <> "" Then
                            sqlStr4 = ""
                            sqlStr4 = sqlStr4 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr4 = sqlStr4 & " VALUES('" & txtBarScale4.Text & "', '" & codeS4V3 & "', " & wordx4(30) & ", getdate(), " & coub4 & ", 4, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr4, "")
                        End If
                    End If

                    'Valve4
                    csql4 = ""
                    csql4 = csql4 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale4.Text & "') and (BOM_RM_Code = '" & codeS4V4 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub4 & ") and (BOM_RM_Scales = 4) "
                    If checkSaveDouble(csql4) < 0 Then
                        If txtBarScale4.Text <> "" And codeS4V4 <> "" Then
                            sqlStr4 = ""
                            sqlStr4 = sqlStr4 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr4 = sqlStr4 & " VALUES('" & txtBarScale4.Text & "', '" & codeS4V4 & "', " & wordx4(38) & ", getdate(), " & coub4 & ", 4, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr4, "")
                        End If
                    End If

                    'Valve5
                    csql5 = ""
                    csql5 = csql5 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale4.Text & "') and (BOM_RM_Code = '" & codeS4V5 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub5 & ") and (BOM_RM_Scales = 4) "
                    If checkSaveDouble(csql5) < 0 Then
                        If txtBarScale4.Text <> "" And codeS4V5 <> "" Then
                            sqlStr4 = ""
                            sqlStr4 = sqlStr4 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                            sqlStr4 = sqlStr4 & " VALUES('" & txtBarScale4.Text & "', '" & codeS4V5 & "', " & wordx4(46) & ", getdate(), " & coub4 & ", 4, getdate()) "
                            Call dbTools.dbSaveSQLsrv(sqlStr4, "")
                        End If
                    End If

                    ShowDataToGridView()
                Catch ex As Exception

                End Try

                Exit Select
            Case 5
                Try


                    txtBarScale5.Text = getDocNo5(txtBarScale5.Text, 5)
                    'Scale#5
                    'Valve1
                    csql1 = ""
                    csql1 = csql1 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale5.Text & "') and (BOM_RM_Code = '" & codeS5V6 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub5 & ") and (BOM_RM_Scales = 5) "
                    If checkSaveDouble(csql1) < 0 Then
                        If txtBarScale5.Text <> "" And codeS5V6 <> "" And (Val(wordx5(14)) > 0 And Val(wordx5(14)) < 1000) Then
                            If wordx5(14) > 3.0 Then
                                sqlStr5 = ""
                                sqlStr5 = sqlStr5 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                                sqlStr5 = sqlStr5 & " VALUES('" & txtBarScale5.Text & "', '" & codeS5V6 & "', " & wordx5(14) & ", getdate(), " & coub5 & ", 5, getdate()) "
                                Call dbTools.dbSaveSQLsrv(sqlStr5, "")
                            End If
                        End If
                    End If

                    ShowDataToGridView()
                Catch ex As Exception

                End Try
                Exit Select
            Case 6
                Try
                    'Scale#6
                    'Valve1
                    csql1 = ""
                    csql1 = csql1 & "Select * From BOMmastF Where (BOM_No = '" & txtBarScale6.Text & "') and (BOM_RM_Code = '" & codeS6V6 & "') and (BOM_RM_Update = convert(varchar(10), getdate(), 102)) and (BOM_RM_Number = " & coub6 & ") and (BOM_RM_Scales = 6) "
                    If checkSaveDouble(csql1) < 0 Then
                        If txtBarScale6.Text <> "" And codeS6V6 <> "" Then
                            If wordx6(14) > 3.0 Then
                                sqlStr6 = ""
                                sqlStr6 = sqlStr6 & "INSERT INTO BOMmastF (BOM_No, BOM_RM_Code, BOM_RM_Values, BOM_RM_Update, BOM_RM_Number, BOM_RM_Scales, BOM_RM_Updatetime) "
                                sqlStr6 = sqlStr6 & " VALUES('" & txtBarScale6.Text & "', '" & codeS6V6 & "', " & wordx6(14) & ", getdate(), " & coub6 & ", 6, getdate()) "
                                Call dbTools.dbSaveSQLsrv(sqlStr6, "")
                            End If
                        End If
                    End If

                    ShowDataToGridView()
                Catch ex As Exception

                End Try
                Exit Select
        End Select
    End Sub
    ' ------------------------------------------------------------------------
    ' Event for connect IP address
    ' ------------------------------------------------------------------------

    Private Sub ConnectIPAddress(ByVal _txtIP As String)
        Try
            MBMaster = New Master(_txtIP, 502)
            AddHandler MBMaster.OnResponseData, New ModbusTCP.Master.ResponseData(AddressOf MBmaster_OnResponseData)
            AddHandler MBMaster.OnException, New ModbusTCP.Master.ExceptionData(AddressOf MBmaster_OnException)
        Catch [error] As SystemException
            'MessageBox.Show([error].Message)
        End Try
    End Sub
    ' ------------------------------------------------------------------------
    ' Event for response data
    ' ------------------------------------------------------------------------
    Private Sub MBmaster_OnResponseData(ByVal ID As UShort, ByVal unit As Byte, ByVal [function] As Byte, ByVal values As Byte())

        If Me.InvokeRequired Then
            Me.BeginInvoke(New Master.ResponseData(AddressOf MBmaster_OnResponseData), New Object() {ID, unit, [function], values})
            Return
        End If

        If stScale1 = 1 And activeScale1 = 1 Then
            data1 = values
            ShowAsW1(Nothing, Nothing)
            InputData1()
        End If

        If stScale2 = 1 And activeScale2 = 1 Then
            data2 = values
            ShowAsW2(Nothing, Nothing)
            InputData2()
        End If

        If stScale3 = 1 And activeScale3 = 1 Then
            data3 = values
            ShowAsW3(Nothing, Nothing)
            InputData3()
        End If

        If stScale4 = 1 And activeScale4 = 1 Then
            data4 = values
            ShowAsW4(Nothing, Nothing)
            InputData4()
        End If

        If stScale5 = 1 And activeScale5 = 1 Then
            data5 = values
            ShowAsW5(Nothing, Nothing)
            InputData5()
        End If

        If stScale6 = 1 And activeScale6 = 1 Then
            data6 = values
            ShowAsW6(Nothing, Nothing)
            InputData6()
        End If

    End Sub
    ' ------------------------------------------------------------------------
    ' Modbus TCP slave exception
    ' ------------------------------------------------------------------------
    Private Sub MBmaster_OnException(ByVal id As UShort, ByVal unit As Byte, ByVal [function] As Byte, ByVal exception As Byte)
        Dim exc As String = "Modbus says error: "
        Select Case exception
            Case Master.excIllegalFunction
                exc += "Illegal function!"
                Exit Select
            Case Master.excIllegalDataAdr
                exc += "Illegal data adress!"
                Exit Select
            Case Master.excIllegalDataVal
                exc += "Illegal data value!"
                Exit Select
            Case Master.excSlaveDeviceFailure
                exc += "Slave device failure!"
                Exit Select
            Case Master.excAck
                exc += "Acknoledge!"
                Exit Select
            Case Master.excGatePathUnavailable
                exc += "Gateway path unavailbale!"
                Exit Select
            Case Master.excExceptionTimeout
                exc += "Slave timed out!"
                Exit Select
            Case Master.excExceptionConnectionLost
                exc += "Connection is lost!"
                Exit Select
            Case Master.excExceptionNotConnected
                exc += "Not connected!"
                Exit Select
        End Select

        MessageBox.Show(exc, "Modbus slave exception")
    End Sub

    Private Sub ShowAsW1(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim word As Integer() = New Integer(0) {}
        'Dim wordx As Double() = New Double(0.0) {}
        'Dim wordxx As Integer = 65536
        If data1.Length < 2 Then
            Return
        End If

        word1 = New Integer(data1.Length \ 2 - 1) {}
        wordx1 = New Double(data1.Length \ 2 - 1) {}

        Dim x As Integer = 0
        While x < data1.Length
            word1(x \ 2) = data1(x) * 256 + data1(x + 1)
            wordx1(x \ 2) = 0.0
            x = x + 2
        End While

        For Each ctrl As Control In grpData1.Controls
            If TypeOf ctrl Is TextBox Then

                Dim xx As Integer = Convert.ToInt16(ctrl.Tag)
                If xx <= data1.GetUpperBound(0) Then
                    ctrl.Text = data1(xx).ToString()
                    ctrl.Visible = True
                Else
                    ctrl.Text = ""
                End If

            End If
        Next
    End Sub

    Private Sub ShowAsW2(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim word As Integer() = New Integer(0) {}
        'Dim wordx As Double() = New Double(0.0) {}
        'Dim wordxx As Integer = 65536
        If data2.Length < 2 Then
            Return
        End If

        word2 = New Integer(data2.Length \ 2 - 1) {}
        wordx2 = New Double(data2.Length \ 2 - 1) {}

        Dim x As Integer = 0
        While x < data2.Length
            word2(x \ 2) = data2(x) * 256 + data2(x + 1)
            wordx2(x \ 2) = 0.0
            x = x + 2
        End While

        For Each ctrl As Control In grpData2.Controls
            If TypeOf ctrl Is TextBox Then

                Dim xx As Integer = Convert.ToInt16(ctrl.Tag)
                If xx <= data2.GetUpperBound(0) Then
                    ctrl.Text = data2(xx).ToString()
                    ctrl.Visible = True
                Else
                    ctrl.Text = ""
                End If

            End If
        Next
    End Sub

    Private Sub ShowAsW3(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim word As Integer() = New Integer(0) {}
        'Dim wordx As Double() = New Double(0.0) {}
        'Dim wordxx As Integer = 65536
        If data3.Length < 2 Then
            Return
        End If

        word3 = New Integer(data3.Length \ 2 - 1) {}
        wordx3 = New Double(data3.Length \ 2 - 1) {}

        Dim x As Integer = 0
        While x < data3.Length
            word3(x \ 2) = data3(x) * 256 + data3(x + 1)
            wordx3(x \ 2) = 0.0
            x = x + 2
        End While

        For Each ctrl As Control In grpData3.Controls
            If TypeOf ctrl Is TextBox Then

                Dim xx As Integer = Convert.ToInt16(ctrl.Tag)
                If xx <= data3.GetUpperBound(0) Then
                    ctrl.Text = data3(xx).ToString()
                    ctrl.Visible = True
                Else
                    ctrl.Text = ""
                End If

            End If
        Next
    End Sub

    Private Sub ShowAsW4(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim word As Integer() = New Integer(0) {}
        'Dim wordx As Double() = New Double(0.0) {}
        'Dim wordxx As Integer = 65536
        If data4.Length < 2 Then
            Return
        End If

        word4 = New Integer(data4.Length \ 2 - 1) {}
        wordx4 = New Double(data4.Length \ 2 - 1) {}

        Dim x As Integer = 0
        While x < data4.Length
            word4(x \ 2) = data4(x) * 256 + data4(x + 1)
            wordx4(x \ 2) = 0.0
            x = x + 2
        End While

        For Each ctrl As Control In grpData4.Controls
            If TypeOf ctrl Is TextBox Then

                Dim xx As Integer = Convert.ToInt16(ctrl.Tag)
                If xx <= data4.GetUpperBound(0) Then
                    ctrl.Text = data4(xx).ToString()
                    ctrl.Visible = True
                Else
                    ctrl.Text = ""
                End If

            End If
        Next
    End Sub

    Private Sub ShowAsW5(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim word As Integer() = New Integer(0) {}
        'Dim wordx As Double() = New Double(0.0) {}
        'Dim wordxx As Integer = 65536
        If data5.Length < 2 Then
            Return
        End If

        word5 = New Integer(data5.Length \ 2 - 1) {}
        wordx5 = New Double(data5.Length \ 2 - 1) {}

        Dim x As Integer = 0
        While x < data5.Length
            word5(x \ 2) = data5(x) * 256 + data5(x + 1)
            wordx5(x \ 2) = 0.0
            x = x + 2
        End While

        For Each ctrl As Control In grpData5.Controls
            If TypeOf ctrl Is TextBox Then

                Dim xx As Integer = Convert.ToInt16(ctrl.Tag)
                If xx <= data5.GetUpperBound(0) Then
                    ctrl.Text = data5(xx).ToString()
                    ctrl.Visible = True
                Else
                    ctrl.Text = ""
                End If

            End If
        Next
    End Sub

    Private Sub ShowAsW6(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim word As Integer() = New Integer(0) {}
        'Dim wordx As Double() = New Double(0.0) {}
        'Dim wordxx As Integer = 65536
        If data6.Length < 2 Then
            Return
        End If

        word6 = New Integer(data6.Length \ 2 - 1) {}
        wordx6 = New Double(data6.Length \ 2 - 1) {}

        Dim x As Integer = 0
        While x < data6.Length
            word6(x \ 2) = data6(x) * 256 + data6(x + 1)
            wordx6(x \ 2) = 0.0
            x = x + 2
        End While

        For Each ctrl As Control In grpData6.Controls
            If TypeOf ctrl Is TextBox Then

                Dim xx As Integer = Convert.ToInt16(ctrl.Tag)
                If xx <= data6.GetUpperBound(0) Then
                    ctrl.Text = data6(xx).ToString()
                    ctrl.Visible = True
                Else
                    ctrl.Text = ""
                End If

            End If
        Next
    End Sub

    ' ------------------------------------------------------------------------
    ' Read values from textboxes
    ' ------------------------------------------------------------------------
    'Private Function GetData(ByVal num As Integer) As Byte()
    '    Dim bits As Boolean() = New Boolean(num - 1) {}
    '    Dim data As Byte() = New [Byte](num - 1) {}
    '    Dim word As Integer() = New Integer(num - 1) {}

    '    ' ------------------------------------------------------------------------
    '    ' Convert data from text boxes

    '    For Each ctrl As Control In grpData.Controls
    '        If TypeOf ctrl Is TextBox Then
    '            Dim x As Integer = Convert.ToInt16(ctrl.Tag)
    '            If (x <= data.GetUpperBound(0)) AndAlso (ctrl.Text <> "") Then
    '                Try
    '                    word(x) = Convert.ToInt16(ctrl.Text)
    '                Catch generatedExceptionName As SystemException
    '                    word(x) = Convert.ToInt16(ctrl.Text)
    '                End Try
    '            Else
    '                Exit For
    '            End If
    '        End If
    '    Next

    '    data = New [Byte](num * 2 - 1) {}
    '    For x As Integer = 0 To num - 1
    '        Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(word(x)))))
    '        data(x * 2) = dat(0)
    '        data(x * 2 + 1) = dat(1)
    '    Next
    '    Return data
    'End Function

    Private Function GetDataScale(ByVal num As Integer, ByVal nscale As Integer) As Byte()
        Dim bits As Boolean() = New Boolean(num - 1) {}
        Dim datas As Byte() = New [Byte](num * 2 - 1) {}

        If nscale = 1 Then data1 = New [Byte](num * 2 - 1) {}
        If nscale = 2 Then data2 = New [Byte](num * 2 - 1) {}
        If nscale = 3 Then data3 = New [Byte](num * 2 - 1) {}
        If nscale = 4 Then data4 = New [Byte](num * 2 - 1) {}
        If nscale = 5 Then data5 = New [Byte](num * 2 - 1) {}
        If nscale = 6 Then data6 = New [Byte](num * 2 - 1) {}

        For x As Integer = 0 To num - 1

            If nscale = 1 Then
                Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(words1(x)))))
                data1(x * 2) = dat(0)
                data1(x * 2 + 1) = dat(1)
            End If
            If nscale = 2 Then
                Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(words2(x)))))
                data2(x * 2) = dat(0)
                data2(x * 2 + 1) = dat(1)
            End If
            If nscale = 3 Then
                Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(words3(x)))))
                data3(x * 2) = dat(0)
                data3(x * 2 + 1) = dat(1)
            End If
            If nscale = 4 Then
                Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(words4(x)))))
                data4(x * 2) = dat(0)
                data4(x * 2 + 1) = dat(1)
            End If
            If nscale = 5 Then
                Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(words5(x)))))
                data5(x * 2) = dat(0)
                data5(x * 2 + 1) = dat(1)
            End If
            If nscale = 6 Then
                Dim dat As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(words6(x)))))
                data6(x * 2) = dat(0)
                data6(x * 2 + 1) = dat(1)
            End If

        Next

        If nscale = 1 Then datas = data1
        If nscale = 2 Then datas = data2
        If nscale = 3 Then datas = data3
        If nscale = 4 Then datas = data4
        If nscale = 5 Then datas = data5
        If nscale = 6 Then datas = data6

        Return datas
    End Function

    Private Sub ReadModbusFromReadHolding()
        Dim ID As UShort = 4
        Dim unit As Byte = Convert.ToByte(txtUnit.Text)
        Dim StartAddress As UShort = ReadStartAdr()
        Dim Length As Byte = Convert.ToByte(txtSize.Text)

        MBMaster.ReadHoldingRegister(ID, unit, StartAddress, Length)
    End Sub

    Private Sub readModbusFromReadInputRegister()
        Dim ID As UShort = 4
        Dim unit As Byte = Convert.ToByte(txtUnit.Text)
        Dim StartAddress As UShort = ReadStartAdr()
        Dim Length As Byte = Convert.ToByte(txtSize.Text)

        MBMaster.ReadInputRegister(ID, unit, StartAddress, Length)
    End Sub

    Private Sub ReadModbusFromReadCoils()
        Dim ID As UShort = 1
        Dim unit As Byte = Convert.ToByte(txtUnit.Text)
        Dim StartAddress As UShort = ReadStartAdr()
        Dim Length As Byte = Convert.ToByte(txtSize.Text)

        MBMaster.ReadCoils(ID, unit, StartAddress, Length)
    End Sub

    Private Function ReadStartAdr() As UShort
        If txtStartAdress.Text.IndexOf("0x", 0, txtStartAdress.Text.Length) = 0 Then
            Dim str As String = txtStartAdress.Text.Replace("0x", "")
            Dim hex As UShort = Convert.ToUInt16(str, 16)
            Return hex
        Else
            Return Convert.ToUInt16(txtStartAdress.Text)
        End If
    End Function

    Private Function ReadAnyStartAdr(ByVal ntStartAdress As Integer) As UShort
        Dim tStartAdress As New TextBox
        tStartAdress.Text = ntStartAdress
        If tStartAdress.Text.IndexOf("0x", 0, tStartAdress.Text.Length) = 0 Then
            Dim str As String = tStartAdress.Text.Replace("0x", "")
            Dim hex As UShort = Convert.ToUInt16(str, 16)
            Return hex
        Else
            Return Convert.ToUInt16(tStartAdress.Text)
        End If
    End Function

    ' ------------------------------------------------------------------------
    ' Button write single coil
    ' ------------------------------------------------------------------------
    'Private Sub WriteSingleCoil()
    '    Dim ID As UShort = 5
    '    Dim unit As Byte = Convert.ToByte(txtUnit.Text)
    '    Dim StartAddress As UShort = ReadStartAdr()

    '    data = GetData(1)
    '    txtSize.Text = "1"

    '    MBMaster.WriteSingleCoils(ID, unit, StartAddress, Convert.ToBoolean(data(0)))
    'End Sub

    ' ------------------------------------------------------------------------
    ' Button write multiple coils
    ' ------------------------------------------------------------------------	
    'Private Sub WriteMultipleCoils()
    '    Dim ID As UShort = 6
    '    Dim unit As Byte = Convert.ToByte(txtUnit.Text)
    '    Dim StartAddress As UShort = ReadStartAdr()

    '    data = GetData(Convert.ToByte(txtSize.Text))
    '    MBMaster.WriteMultipleCoils(ID, unit, StartAddress, Convert.ToByte(txtSize.Text), data)
    'End Sub

    ' ------------------------------------------------------------------------
    ' Button write single register
    ' ------------------------------------------------------------------------
    'Private Sub WriteSingleReg()
    '    Dim ID As UShort = 7
    '    Dim unit As Byte = Convert.ToByte(txtUnit.Text)
    '    Dim StartAddress As UShort = ReadStartAdr()

    '    data = GetData(2)
    '    txtSize.Text = "1"
    '    txtdata.Text = data(0).ToString()

    '    MBMaster.WriteSingleRegister(ID, unit, StartAddress, data)
    'End Sub

    ' ------------------------------------------------------------------------
    ' Button write multiple register
    ' ------------------------------------------------------------------------	
    'Private Sub WriteMultipleReg()
    '    Dim ID As UShort = 8
    '    Dim unit As Byte = Convert.ToByte(txtUnit.Text)
    '    Dim StartAddress As UShort = ReadStartAdr()

    '    data = GetData(Convert.ToByte(txtSize.Text))
    '    MBMaster.WriteMultipleRegister(ID, unit, StartAddress, data)
    'End Sub

    Private Sub InputData1()
        Try
            nc1 += 1
            If nc1 = 1000000 Then nc1 = 0
            hexCheck1.Text = data1(0)
            hexCheck2.Text = data1(1)
            hexCheck3.Text = data1(2)
            hexCheck4.Text = data1(3)
            hexCheck5.Text = data1(4)
            hexCheck6.Text = data1(5)
            hexCheck7.Text = data1(6)
            hexCheck8.Text = data1(7)
            hexCheck9.Text = data1(8)
            hexCheck10.Text = data1(9)
            '-------- Read values finish batch --------
            wordx1(14) = (word1(14) + (word1(15) * wordxx)) / 10
            wordx1(22) = (word1(22) + (word1(23) * wordxx)) / 10
            wordx1(30) = (word1(30) + (word1(31) * wordxx)) / 10
            wordx1(38) = (word1(38) + (word1(39) * wordxx)) / 10
            wordx1(46) = (word1(46) + (word1(47) * wordxx)) / 10
            wordx1(54) = (word1(54) + (word1(55) * wordxx)) / 10

            wordx1(58) = (word1(58) + (word1(59) * wordxx))
            wordx1(60) = (word1(60) + (word1(61) * wordxx))
            wordx1(62) = (word1(62) + (word1(63) * wordxx))
            wordx1(105) = word1(105)
            wordx1(106) = word1(106)
            '------- Convert Barcode --------
            If wordx1(105) = 1 Or stDoScale1 = 0 Then
                txtBarScale1.Text = ConvertHexToDec(Val(hexCheck2.Text)) & ConvertHexToDec(Val(hexCheck1.Text)) &
            ConvertHexToDec(Val(hexCheck4.Text)) & ConvertHexToDec(Val(hexCheck3.Text)) &
            ConvertHexToDec(Val(hexCheck6.Text)) & ConvertHexToDec(Val(hexCheck5.Text)) &
            ConvertHexToDec(Val(hexCheck8.Text)) & ConvertHexToDec(Val(hexCheck7.Text)) &
            ConvertHexToDec(Val(hexCheck10.Text)) & ConvertHexToDec(Val(hexCheck9.Text))
            End If

        Catch ex As Exception

        End Try

        Try
            '-------- Read SQL from database --------
            If wordx1(105) = 1 Then
                FillNullwords(1)
                Call runFindScale(1)
                If nrec1 > 0 Then
                    words1(105) = 0
                    words1(106) = 1

                    '-------- Write values to scale --------
                    Dim cID As UShort = 8
                    Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                    Dim cStartAddress As UShort = ReadStartAdr()
                    txtIP.Text = _svIPAddress1
                    ConnectIPAddress(txtIP.Text)

                    data1 = GetDataScale(Convert.ToByte(txtSize.Text), 1)
                    MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data1)
                    'Else
                    '    lblError.Text = "เครื่องชั่ง #1 ไม่พบเลขที่ใบคุม " & txtBarScale1.Text & "ในระบบ"
                End If
            Else
                If stDoScale1 = 0 Then
                    FillNullwords(1)
                    Call runFindScale(1)
                    stDoScale1 = 1
                End If

            End If

            '-------- Check finish batch --------
            chkFScale1 = wordx1(58)
            cstart1 = wordx1(60)
            coub1 = wordx1(62)

            If cstart1 = 0 Then
                btnRunning1.Text = "Wait Run..."
                btnRunning1.ForeColor = Color.Black

                If chkFScale1 = 1 Then
                    '-------- Save weight finish to database --------
                    SaveDataScale(1)
                    '-------- Write stattus to save finish --------
                    Try
                        words1(108) = 1
                        words1(58) = 1
                        words1(62) = coub1
                        Dim cID As UShort = 8
                        Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                        Dim cStartAddress As UShort = ReadStartAdr()
                        txtIP.Text = _svIPAddress1
                        ConnectIPAddress(txtIP.Text)

                        data1 = GetDataScale(Convert.ToByte(txtSize.Text), 1)
                        MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data1)

                        'chkFScale1 = 0
                        'cstart1 = 1
                    Catch ex As Exception

                    End Try
                End If

            Else
                btnRunning1.Text = "Running..."
                btnRunning1.ForeColor = Color.Green
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub InputData2()
        Try
            nc2 += 1
            If nc2 = 1000000 Then nc2 = 0
            hexCheck1.Text = data2(0)
            hexCheck2.Text = data2(1)
            hexCheck3.Text = data2(2)
            hexCheck4.Text = data2(3)
            hexCheck5.Text = data2(4)
            hexCheck6.Text = data2(5)
            hexCheck7.Text = data2(6)
            hexCheck8.Text = data2(7)
            hexCheck9.Text = data2(8)
            hexCheck10.Text = data2(9)
            '-------- Read values finish batch --------
            wordx2(14) = (word2(14) + (word2(15) * wordxx)) / 10
            wordx2(22) = (word2(22) + (word2(23) * wordxx)) / 10
            wordx2(30) = (word2(30) + (word2(31) * wordxx)) / 10
            wordx2(38) = (word2(38) + (word2(39) * wordxx)) / 10
            wordx2(46) = (word2(46) + (word2(47) * wordxx)) / 10
            wordx2(54) = (word2(54) + (word2(55) * wordxx)) / 10

            wordx2(58) = (word2(58) + (word2(59) * wordxx))
            wordx2(60) = (word2(60) + (word2(61) * wordxx))
            wordx2(62) = (word2(62) + (word2(63) * wordxx))
            wordx2(105) = word2(105)
            wordx2(106) = word2(106)
            '------- Convert Barcode --------
            If wordx2(105) = 1 Or stDoScale2 = 0 Then
                txtBarScale2.Text = ConvertHexToDec(Val(hexCheck2.Text)) & ConvertHexToDec(Val(hexCheck1.Text)) &
            ConvertHexToDec(Val(hexCheck4.Text)) & ConvertHexToDec(Val(hexCheck3.Text)) &
            ConvertHexToDec(Val(hexCheck6.Text)) & ConvertHexToDec(Val(hexCheck5.Text)) &
            ConvertHexToDec(Val(hexCheck8.Text)) & ConvertHexToDec(Val(hexCheck7.Text)) &
            ConvertHexToDec(Val(hexCheck10.Text)) & ConvertHexToDec(Val(hexCheck9.Text))
            End If

        Catch ex As Exception

        End Try

        Try
            '-------- Read SQL from database --------
            If wordx2(105) = 1 Then
                FillNullwords(2)
                Call runFindScale(2)
                If nrec2 > 0 Then
                    words2(105) = 0
                    words2(106) = 1

                    '-------- Write values to scale --------
                    Dim cID As UShort = 8
                    Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                    Dim cStartAddress As UShort = ReadStartAdr()
                    txtIP.Text = _svIPAddress2
                    ConnectIPAddress(txtIP.Text)

                    data2 = GetDataScale(Convert.ToByte(txtSize.Text), 2)
                    MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data2)
                    'Else
                    '    lblError.Text = "เครื่องชั่ง #2 ไม่พบเลขที่ใบคุม " & txtBarScale2.Text & "ในระบบ"
                End If
            Else
                If stDoScale2 = 0 Then
                    FillNullwords(2)
                    Call runFindScale(2)
                    stDoScale2 = 1
                End If
            End If

            '-------- Check finish batch --------
            chkFScale2 = wordx2(58)
            cstart2 = wordx2(60)
            coub2 = wordx2(62)

            If cstart2 = 0 Then
                btnRunning2.Text = "Wait Run..."
                btnRunning2.ForeColor = Color.Black

                If chkFScale2 = 1 Then
                    '-------- Save weight finish to database --------
                    SaveDataScale(2)
                    '-------- Write stattus to save finish --------
                    Try
                        words2(108) = 1
                        words2(58) = 1
                        words2(62) = coub2
                        words2(105) = 0
                        Dim cID As UShort = 8
                        Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                        Dim cStartAddress As UShort = ReadStartAdr()
                        txtIP.Text = _svIPAddress2
                        ConnectIPAddress(txtIP.Text)

                        data2 = GetDataScale(Convert.ToByte(txtSize.Text), 2)
                        MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data2)

                        'chkFScale1 = 0
                        'cstart1 = 1
                    Catch ex As Exception

                    End Try
                End If

            Else
                btnRunning2.Text = "Running..."
                btnRunning2.ForeColor = Color.Green
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub InputData3()
        Try
            nc3 += 1
            If nc3 = 1000000 Then nc3 = 0
            hexCheck1.Text = data3(0)
            hexCheck2.Text = data3(1)
            hexCheck3.Text = data3(2)
            hexCheck4.Text = data3(3)
            hexCheck5.Text = data3(4)
            hexCheck6.Text = data3(5)
            hexCheck7.Text = data3(6)
            hexCheck8.Text = data3(7)
            hexCheck9.Text = data3(8)
            hexCheck10.Text = data3(9)
            '-------- Read values finish batch --------
            wordx3(14) = (word3(14) + (word3(15) * wordxx)) / 10
            wordx3(22) = (word3(22) + (word3(23) * wordxx)) / 10
            wordx3(30) = (word3(30) + (word3(31) * wordxx)) / 10
            wordx3(38) = (word3(38) + (word3(39) * wordxx)) / 10
            wordx3(46) = (word3(46) + (word3(47) * wordxx)) / 10
            wordx3(54) = (word3(54) + (word3(55) * wordxx)) / 10

            wordx3(58) = (word3(58) + (word3(59) * wordxx))
            wordx3(60) = (word3(60) + (word3(61) * wordxx))
            wordx3(62) = (word3(62) + (word3(63) * wordxx))
            wordx3(105) = word3(105)
            wordx3(106) = word3(106)
            '------- Convert Barcode --------
            If wordx3(105) = 1 Or stDoScale3 = 0 Then
                txtBarScale3.Text = ConvertHexToDec(Val(hexCheck2.Text)) & ConvertHexToDec(Val(hexCheck1.Text)) &
            ConvertHexToDec(Val(hexCheck4.Text)) & ConvertHexToDec(Val(hexCheck3.Text)) &
            ConvertHexToDec(Val(hexCheck6.Text)) & ConvertHexToDec(Val(hexCheck5.Text)) &
            ConvertHexToDec(Val(hexCheck8.Text)) & ConvertHexToDec(Val(hexCheck7.Text)) &
            ConvertHexToDec(Val(hexCheck10.Text)) & ConvertHexToDec(Val(hexCheck9.Text))
            End If

        Catch ex As Exception

        End Try

        Try
            '-------- Read SQL from database --------
            If wordx3(105) = 1 Then
                FillNullwords(3)
                Call runFindScale(3)
                If nrec3 > 0 Then
                    words3(105) = 0
                    words3(106) = 1

                    '-------- Write values to scale --------
                    Dim cID As UShort = 8
                    Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                    Dim cStartAddress As UShort = ReadStartAdr()
                    txtIP.Text = _svIPAddress3
                    ConnectIPAddress(txtIP.Text)

                    data3 = GetDataScale(Convert.ToByte(txtSize.Text), 3)
                    'MsgBox(data3(24) & " : " & data3(32) & " : " & data3(33))
                    MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data3)
                    'Else
                    '    lblError.Text = "เครื่องชั่ง #3 ไม่พบเลขที่ใบคุม " & txtBarScale3.Text & "ในระบบ"
                End If
            Else
                If stDoScale3 = 0 Then
                    FillNullwords(3)
                    Call runFindScale(3)
                    stDoScale3 = 1
                End If
            End If

            '-------- Check finish batch --------
            chkFScale3 = wordx3(58)
            cstart3 = wordx3(60)
            coub3 = wordx3(62)

            If cstart3 = 0 Then
                btnRunning3.Text = "Wait Run..."
                btnRunning3.ForeColor = Color.Black

                If chkFScale3 = 1 Then
                    '-------- Save weight finish to database --------
                    SaveDataScale(3)
                    '-------- Write stattus to save finish --------
                    Try
                        words3(108) = 1
                        words3(58) = 1
                        words3(62) = coub3
                        words3(105) = 0
                        Dim cID As UShort = 8
                        Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                        Dim cStartAddress As UShort = ReadStartAdr()
                        txtIP.Text = _svIPAddress3
                        ConnectIPAddress(txtIP.Text)

                        data3 = GetDataScale(Convert.ToByte(txtSize.Text), 3)
                        MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data3)

                        'chkFScale1 = 0
                        'cstart1 = 1
                    Catch ex As Exception

                    End Try
                End If

            Else
                btnRunning3.Text = "Running..."
                btnRunning3.ForeColor = Color.Green

            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub updateCM_Running()

        txtSQL = "Insert Into cmMixMast "

    End Sub
    Private Sub InputData4()
        Try
            nc4 += 1
            If nc4 = 1000000 Then nc4 = 0
            hexCheck1.Text = data4(0)
            hexCheck2.Text = data4(1)
            hexCheck3.Text = data4(2)
            hexCheck4.Text = data4(3)
            hexCheck5.Text = data4(4)
            hexCheck6.Text = data4(5)
            hexCheck7.Text = data4(6)
            hexCheck8.Text = data4(7)
            hexCheck9.Text = data4(8)
            hexCheck10.Text = data4(9)
            '-------- Read values finish batch --------
            wordx4(14) = (word4(14) + (word4(15) * wordxx)) / 10
            wordx4(22) = (word4(22) + (word4(23) * wordxx)) / 10
            wordx4(30) = (word4(30) + (word4(31) * wordxx)) / 10
            wordx4(38) = (word4(38) + (word4(39) * wordxx)) / 10
            wordx4(46) = (word4(46) + (word4(47) * wordxx)) / 10
            wordx4(54) = (word4(54) + (word4(55) * wordxx)) / 10

            wordx4(58) = (word4(58) + (word4(59) * wordxx))
            wordx4(60) = (word4(60) + (word4(61) * wordxx))
            wordx4(62) = (word4(62) + (word4(63) * wordxx))
            wordx4(105) = word4(105)
            wordx4(106) = word4(106)
            '------- Convert Barcode --------
            If wordx4(105) = 1 Or stDoScale4 = 0 Then
                txtBarScale4.Text = ConvertHexToDec(Val(hexCheck2.Text)) & ConvertHexToDec(Val(hexCheck1.Text)) &
            ConvertHexToDec(Val(hexCheck4.Text)) & ConvertHexToDec(Val(hexCheck3.Text)) &
            ConvertHexToDec(Val(hexCheck6.Text)) & ConvertHexToDec(Val(hexCheck5.Text)) &
            ConvertHexToDec(Val(hexCheck8.Text)) & ConvertHexToDec(Val(hexCheck7.Text)) &
            ConvertHexToDec(Val(hexCheck10.Text)) & ConvertHexToDec(Val(hexCheck9.Text))
            End If

        Catch ex As Exception

        End Try

        Try
            '-------- Read SQL from database --------
            If wordx4(105) = 1 Then
                FillNullwords(4)
                Call runFindScale(4)
                If nrec4 > 0 Then
                    words4(105) = 0
                    words4(106) = 1

                    '-------- Write values to scale --------
                    Dim cID As UShort = 8
                    Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                    Dim cStartAddress As UShort = ReadStartAdr()
                    txtIP.Text = _svIPAddress4
                    ConnectIPAddress(txtIP.Text)

                    data4 = GetDataScale(Convert.ToByte(txtSize.Text), 4)
                    MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data4)
                Else
                    lblError.Text = "เครื่องชั่ง #4 ไม่พบเลขที่ใบคุม " & txtBarScale4.Text & "ในระบบ"
                End If
            Else
                If stDoScale4 = 0 Then
                    FillNullwords(4)
                    Call runFindScale(4)
                    stDoScale4 = 1
                End If
            End If

            '-------- Check finish batch --------
            chkFScale4 = wordx4(58)
            cstart4 = wordx4(60)
            coub4 = wordx4(62)

            If cstart4 = 0 Then
                btnRunning4.Text = "Wait Run..."
                btnRunning4.ForeColor = Color.Black

                If chkFScale4 = 1 Then
                    '-------- Save weight finish to database --------
                    SaveDataScale(4)
                    '-------- Write stattus to save finish --------
                    Try
                        words4(108) = 1
                        words4(58) = 1
                        words4(62) = coub4
                        words4(105) = 0
                        Dim cID As UShort = 8
                        Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                        Dim cStartAddress As UShort = ReadStartAdr()
                        txtIP.Text = _svIPAddress4
                        ConnectIPAddress(txtIP.Text)

                        data4 = GetDataScale(Convert.ToByte(txtSize.Text), 4)
                        MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data4)

                        'chkFScale1 = 0
                        'cstart1 = 1
                    Catch ex As Exception

                    End Try
                End If

            Else
                btnRunning4.Text = "Running..."
                btnRunning4.ForeColor = Color.Green
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub InputData5()
        Try
            nc5 += 1
            If nc5 = 1000000 Then nc5 = 0
            hexCheck1.Text = data5(0)
            hexCheck2.Text = data5(1)
            hexCheck3.Text = data5(2)
            hexCheck4.Text = data5(3)
            hexCheck5.Text = data5(4)
            hexCheck6.Text = data5(5)
            hexCheck7.Text = data5(6)
            hexCheck8.Text = data5(7)
            hexCheck9.Text = data5(8)
            hexCheck10.Text = data5(9)
            '-------- Read values finish batch --------
            wordx5(14) = (word5(14) + (word5(15) * wordxx)) / 10
            wordx5(22) = (word5(22) + (word5(23) * wordxx)) / 10
            wordx5(30) = (word5(30) + (word5(31) * wordxx)) / 10
            wordx5(38) = (word5(38) + (word5(39) * wordxx)) / 10
            wordx5(46) = (word5(46) + (word5(47) * wordxx)) / 10
            wordx5(54) = (word5(54) + (word5(55) * wordxx)) / 10

            wordx5(58) = (word5(58) + (word5(59) * wordxx))
            wordx5(60) = (word5(60) + (word5(61) * wordxx))
            wordx5(62) = (word5(62) + (word5(63) * wordxx))
            wordx5(105) = word5(105)
            wordx5(106) = word5(106)
            wordx5(108) = word5(108)
            'MsgBox(ConvertHexToDec(Val(hexCheck2.Text)))
            '------- Convert Barcode --------
            If wordx5(105) = 1 Or stDoScale5 = 0 Then
                txtBarScale5.Text = ConvertHexToDec(Val(hexCheck2.Text)) & ConvertHexToDec(Val(hexCheck1.Text)) &
            ConvertHexToDec(Val(hexCheck4.Text)) & ConvertHexToDec(Val(hexCheck3.Text)) &
            ConvertHexToDec(Val(hexCheck6.Text)) & ConvertHexToDec(Val(hexCheck5.Text)) &
            ConvertHexToDec(Val(hexCheck8.Text)) & ConvertHexToDec(Val(hexCheck7.Text)) &
            ConvertHexToDec(Val(hexCheck10.Text)) & ConvertHexToDec(Val(hexCheck9.Text))
            End If

        Catch ex As Exception

        End Try

        Try
            '-------- Read SQL from database --------
            If wordx5(105) = 1 Then
                FillNullwords(5)
                Call runFindScale(5)
                If nrec5 > 0 Then
                    words5(105) = 0
                    words5(106) = 1

                    '-------- Write values to scale --------
                    Dim cID As UShort = 8
                    Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                    Dim cStartAddress As UShort = ReadStartAdr()
                    txtIP.Text = _svIPAddress5
                    ConnectIPAddress(txtIP.Text)

                    data5 = GetDataScale(Convert.ToByte(txtSize.Text), 5)
                    MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data5)
                    'Else
                    '    lblError.Text = "เครื่องชั่ง #5 ไม่พบเลขที่ใบคุม " & txtBarScale5.Text & "ในระบบ"
                End If
            Else
                If stDoScale5 = 0 Then
                    FillNullwords(5)
                    Call runFindScale(5)
                    stDoScale5 = 1
                End If
            End If

            '-------- Check finish batch --------
            'MsgBox(wordx5(58))
            chkFScale5 = wordx5(58)
            cstart5 = wordx5(60)
            coub5 = wordx5(62)

            If cstart5 = 1 Then
                btnRunning5.Text = "Running..."
                btnRunning5.ForeColor = Color.Green

                If chkFScale5 = 1 Then
                    '-------- Save weight finish to database --------
                    SaveDataScale(5)
                    '-------- Write stattus to save finish --------
                    Try
                        words5(108) = 1
                        words5(58) = 1
                        words5(14) = wordx5(14)
                        words5(62) = coub5
                        words5(105) = 0
                        Dim cID As UShort = 8
                        Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                        Dim cStartAddress As UShort = ReadStartAdr()
                        txtIP.Text = _svIPAddress5
                        ConnectIPAddress(txtIP.Text)

                        data5 = GetDataScale(Convert.ToByte(txtSize.Text), 5)
                        MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data5)
                        words5(14) = 0

                    Catch ex As Exception

                    End Try
                End If

            Else
                btnRunning5.Text = "Wait Run..."
                btnRunning5.ForeColor = Color.Black
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub InputData6()
        Try
            nc6 += 1
            If nc6 = 1000000 Then nc6 = 0
            hexCheck1.Text = data6(0)
            hexCheck2.Text = data6(1)
            hexCheck3.Text = data6(2)
            hexCheck4.Text = data6(3)
            hexCheck5.Text = data6(4)
            hexCheck6.Text = data6(5)
            hexCheck7.Text = data6(6)
            hexCheck8.Text = data6(7)
            hexCheck9.Text = data6(8)
            hexCheck10.Text = data6(9)
            '-------- Read values finish batch --------
            wordx6(14) = (word6(14) + (word6(15) * wordxx)) / 10
            wordx6(22) = (word6(22) + (word6(23) * wordxx)) / 10
            wordx6(30) = (word6(30) + (word6(31) * wordxx)) / 10
            wordx6(38) = (word6(38) + (word6(39) * wordxx)) / 10
            wordx6(46) = (word6(46) + (word6(47) * wordxx)) / 10
            wordx6(54) = (word6(54) + (word6(55) * wordxx)) / 10

            wordx6(58) = (word6(58) + (word6(59) * wordxx))
            wordx6(60) = (word6(60) + (word6(61) * wordxx))
            wordx6(62) = (word1(62) + (word1(63) * wordxx))
            wordx6(105) = word6(105)
            wordx6(106) = word6(106)
            '------- Convert Barcode --------
            If wordx6(105) = 1 Or stDoScale6 = 0 Then
                txtBarScale6.Text = ConvertHexToDec(Val(hexCheck2.Text)) & ConvertHexToDec(Val(hexCheck1.Text)) &
            ConvertHexToDec(Val(hexCheck4.Text)) & ConvertHexToDec(Val(hexCheck3.Text)) &
            ConvertHexToDec(Val(hexCheck6.Text)) & ConvertHexToDec(Val(hexCheck5.Text)) &
            ConvertHexToDec(Val(hexCheck8.Text)) & ConvertHexToDec(Val(hexCheck7.Text)) &
            ConvertHexToDec(Val(hexCheck10.Text)) & ConvertHexToDec(Val(hexCheck9.Text))
            End If

        Catch ex As Exception

        End Try

        Try
            '-------- Read SQL from database --------
            If wordx6(105) = 1 Then
                FillNullwords(6)
                Call runFindScale(6)
                If nrec6 > 0 Then
                    words6(105) = 0
                    words6(106) = 1

                    '-------- Write values to scale --------
                    Dim cID As UShort = 8
                    Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                    Dim cStartAddress As UShort = ReadStartAdr()
                    txtIP.Text = _svIPAddress6
                    ConnectIPAddress(txtIP.Text)

                    data6 = GetDataScale(Convert.ToByte(txtSize.Text), 6)
                    MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data6)
                    'Else
                    '    lblError.Text = "เครื่องชั่ง #6 ไม่พบเลขที่ใบคุม " & txtBarScale6.Text & "ในระบบ"
                End If
            Else
                If stDoScale6 = 0 Then
                    FillNullwords(6)
                    Call runFindScale(6)
                    stDoScale6 = 1
                End If
            End If

            '-------- Check finish batch --------
            chkFScale6 = wordx6(58)
            cstart6 = wordx6(60)
            coub6 = wordx6(62)
            If cstart6 = 0 Then
                btnRunning6.Text = "Wait Run..."
                btnRunning6.ForeColor = Color.Black

                If chkFScale6 = 1 Then
                    '-------- Save weight finish to database --------
                    SaveDataScale(6)
                    '-------- Write stattus to save finish --------
                    Try
                        words6(108) = 1
                        words6(58) = 1
                        words6(62) = coub6
                        words6(105) = 0
                        Dim cID As UShort = 8
                        Dim cunit As Byte = Convert.ToByte(txtUnit.Text)
                        Dim cStartAddress As UShort = ReadStartAdr()
                        txtIP.Text = _svIPAddress6
                        ConnectIPAddress(txtIP.Text)

                        data6 = GetDataScale(Convert.ToByte(txtSize.Text), 6)
                        MBMaster.WriteMultipleRegister(cID, cunit, cStartAddress, data6)

                        'chkFScale1 = 0
                        'cstart1 = 1
                    Catch ex As Exception

                    End Try
                End If

            Else
                btnRunning6.Text = "Running..."
                btnRunning6.ForeColor = Color.Green
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tmrConnectIP1_Tick(sender As Object, e As EventArgs) Handles tmrConnectIP1.Tick
        TextBox1.Text = "1"
        If stScale1 = 0 Then btnConScale1_Click(Nothing, Nothing)

        If stScale1 = 1 Then
            activeScale1 = 1
            activeScale2 = 0
            activeScale3 = 0
            activeScale4 = 0
            activeScale5 = 0
            activeScale6 = 0

            txtIP.Text = _svIPAddress1
            ConnectIPAddress(txtIP.Text)
            readModbusFromReadInputRegister()
        End If

        tmrConnectIP1.Stop()
        tmrConnectIP2.Start()
        tmrConnectIP3.Stop()
        tmrConnectIP4.Stop()
        tmrConnectIP5.Stop()
        tmrConnectIP6.Stop()
    End Sub

    Private Sub tmrConnectIP2_Tick(sender As Object, e As EventArgs) Handles tmrConnectIP2.Tick
        TextBox1.Text = "2"
        If stScale2 = 0 Then btnConScale2_Click(Nothing, Nothing)

        If stScale2 = 1 Then
            activeScale1 = 0
            activeScale2 = 1
            activeScale3 = 0
            activeScale4 = 0
            activeScale5 = 0
            activeScale6 = 0

            txtIP.Text = _svIPAddress2
            ConnectIPAddress(txtIP.Text)
            readModbusFromReadInputRegister()
        End If

        tmrConnectIP1.Stop()
        tmrConnectIP2.Stop()
        tmrConnectIP3.Start()
        tmrConnectIP4.Stop()
        tmrConnectIP5.Stop()
        tmrConnectIP6.Stop()
    End Sub

    Private Sub tmrConnectIP3_Tick(sender As Object, e As EventArgs) Handles tmrConnectIP3.Tick
        TextBox1.Text = "3"
        If stScale3 = 0 Then btnConScale3_Click(Nothing, Nothing)

        If stScale3 = 1 Then
            activeScale1 = 0
            activeScale2 = 0
            activeScale3 = 1
            activeScale4 = 0
            activeScale5 = 0
            activeScale6 = 0

            txtIP.Text = _svIPAddress3
            ConnectIPAddress(txtIP.Text)
            readModbusFromReadInputRegister()
        End If

        tmrConnectIP1.Stop()
        tmrConnectIP2.Stop()
        tmrConnectIP3.Stop()
        tmrConnectIP4.Start()
        tmrConnectIP5.Stop()
        tmrConnectIP6.Stop()
    End Sub

    Private Sub tmrConnectIP4_Tick(sender As Object, e As EventArgs) Handles tmrConnectIP4.Tick
        TextBox1.Text = "4"
        If stScale4 = 0 Then btnConScale4_Click(Nothing, Nothing)

        If stScale4 = 1 Then
            activeScale1 = 0
            activeScale2 = 0
            activeScale3 = 0
            activeScale4 = 1
            activeScale5 = 0
            activeScale6 = 0

            txtIP.Text = _svIPAddress4
            ConnectIPAddress(txtIP.Text)
            readModbusFromReadInputRegister()
        End If

        tmrConnectIP1.Stop()
        tmrConnectIP2.Stop()
        tmrConnectIP3.Stop()
        tmrConnectIP4.Stop()
        tmrConnectIP5.Start()
        tmrConnectIP6.Stop()
    End Sub

    Private Sub tmrConnectIP5_Tick(sender As Object, e As EventArgs) Handles tmrConnectIP5.Tick
        TextBox1.Text = "5"
        If stScale5 = 0 Then btnConScale5_Click(Nothing, Nothing)

        If stScale5 = 1 Then
            activeScale1 = 0
            activeScale2 = 0
            activeScale3 = 0
            activeScale4 = 0
            activeScale5 = 1
            activeScale6 = 0

            txtIP.Text = _svIPAddress5
            ConnectIPAddress(txtIP.Text)
            readModbusFromReadInputRegister()
        End If

        tmrConnectIP1.Stop()
        tmrConnectIP2.Stop()
        tmrConnectIP3.Stop()
        tmrConnectIP4.Stop()
        tmrConnectIP5.Stop()
        tmrConnectIP6.Start()
    End Sub

    Private Sub tmrConnectIP6_Tick(sender As Object, e As EventArgs) Handles tmrConnectIP6.Tick
        TextBox1.Text = "6"
        If stScale6 = 0 Then btnConScale6_Click(Nothing, Nothing)

        If stScale6 = 1 Then
            activeScale1 = 0
            activeScale2 = 0
            activeScale3 = 0
            activeScale4 = 0
            activeScale5 = 0
            activeScale6 = 1

            txtIP.Text = _svIPAddress6
            ConnectIPAddress(txtIP.Text)
            readModbusFromReadInputRegister()
        End If

        tmrConnectIP1.Start()
        tmrConnectIP2.Stop()
        tmrConnectIP3.Stop()
        tmrConnectIP4.Stop()
        tmrConnectIP5.Stop()
        tmrConnectIP6.Stop()
    End Sub

    Private Sub btnConScale1_Click(sender As Object, e As EventArgs) Handles btnConScale1.Click
        If My.Computer.Network.Ping(_svIPAddress1) Then
            btnConScale1.BackColor = Color.Green
            btnConScale1.ForeColor = Color.Red
            btnConScale1.Text = "Connected"
            stScale1 = 1
        Else
            btnConScale1.BackColor = Color.Red
            btnConScale1.ForeColor = Color.White
            btnConScale1.Text = "Connect"
            stScale1 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #1 :: IP Address " & _svIPAddress1 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnConScale2_Click(sender As Object, e As EventArgs) Handles btnConScale2.Click
        If My.Computer.Network.Ping(_svIPAddress2) Then
            btnConScale2.BackColor = Color.Green
            btnConScale2.ForeColor = Color.Red
            btnConScale2.Text = "Connected"
            stScale2 = 1
        Else
            btnConScale2.BackColor = Color.Red
            btnConScale2.ForeColor = Color.White
            btnConScale2.Text = "Connect"
            stScale2 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #2 :: IP Address " & _svIPAddress2 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnConScale3_Click(sender As Object, e As EventArgs) Handles btnConScale3.Click
        If My.Computer.Network.Ping(_svIPAddress3) Then
            btnConScale3.BackColor = Color.Green
            btnConScale3.ForeColor = Color.Red
            btnConScale3.Text = "Connected"
            stScale3 = 1
        Else
            btnConScale3.BackColor = Color.Red
            btnConScale3.ForeColor = Color.White
            btnConScale3.Text = "Connect"
            stScale3 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #3 :: IP Address " & _svIPAddress3 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnConScale4_Click(sender As Object, e As EventArgs) Handles btnConScale4.Click
        If My.Computer.Network.Ping(_svIPAddress4) Then
            btnConScale4.BackColor = Color.Green
            btnConScale4.ForeColor = Color.Red
            btnConScale4.Text = "Connected"
            stScale4 = 1
        Else
            btnConScale4.BackColor = Color.Red
            btnConScale4.ForeColor = Color.White
            btnConScale4.Text = "Connect"
            stScale4 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #4 :: IP Address " & _svIPAddress4 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnConScale5_Click(sender As Object, e As EventArgs) Handles btnConScale5.Click
        If My.Computer.Network.Ping(_svIPAddress5) Then
            btnConScale5.BackColor = Color.Green
            btnConScale5.ForeColor = Color.Red
            btnConScale5.Text = "Connected"
            stScale5 = 1
        Else
            btnConScale5.BackColor = Color.Red
            btnConScale5.ForeColor = Color.White
            btnConScale5.Text = "Connect"
            stScale5 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #5 :: IP Address " & _svIPAddress5 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnConScale6_Click(sender As Object, e As EventArgs) Handles btnConScale6.Click
        If My.Computer.Network.Ping(_svIPAddress6) Then
            btnConScale6.BackColor = Color.Green
            btnConScale6.ForeColor = Color.Red
            btnConScale6.Text = "Connected"
            stScale6 = 1
        Else
            btnConScale6.BackColor = Color.Red
            btnConScale6.ForeColor = Color.White
            btnConScale6.Text = "Connect"
            stScale6 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #6 :: IP Address " & _svIPAddress6 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        dbTools.closeDB()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dbTools.openDB()
        setYear()
        LoadSetting()

        data1 = New Byte(-1) {}
        data2 = New Byte(-1) {}
        data3 = New Byte(-1) {}
        data4 = New Byte(-1) {}
        data5 = New Byte(-1) {}
        data6 = New Byte(-1) {}

        ResizeData1()
        ResizeData2()
        ResizeData3()
        ResizeData4()
        ResizeData5()
        ResizeData6()

        FillNullwords(0)

        tmrConnectIP1.Interval = _Interval_values
        tmrConnectIP2.Interval = _Interval_values
        tmrConnectIP3.Interval = _Interval_values
        tmrConnectIP4.Interval = _Interval_values
        tmrConnectIP5.Interval = _Interval_values
        tmrConnectIP6.Interval = _Interval_values

        txtUnit.Text = _svUnit
        txtStartAdress.Text = _svStartAddress
        txtSize.Text = _svSize
        txtIP.Text = _svIPAddress1

        ShowDataToGridView()


        If My.Computer.Network.Ping(_svIPAddress1) Then
            btnConScale1.BackColor = Color.Green
            btnConScale1.ForeColor = Color.Red
            btnConScale1.Text = "Connected"
            stScale1 = 1
        Else
            btnConScale1.BackColor = Color.Red
            btnConScale1.ForeColor = Color.White
            btnConScale1.Text = "Connect"
            stScale1 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #1 :: IP Address " & _svIPAddress1 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        If My.Computer.Network.Ping(_svIPAddress2) Then
            btnConScale2.BackColor = Color.Green
            btnConScale2.ForeColor = Color.Red
            btnConScale2.Text = "Connected"
            stScale2 = 1
        Else
            btnConScale2.BackColor = Color.Red
            btnConScale2.ForeColor = Color.White
            btnConScale2.Text = "Connect"
            stScale2 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #2 :: IP Address " & _svIPAddress2 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        If My.Computer.Network.Ping(_svIPAddress3) Then
            btnConScale3.BackColor = Color.Green
            btnConScale3.ForeColor = Color.Red
            btnConScale3.Text = "Connected"
            stScale3 = 1
        Else
            btnConScale3.BackColor = Color.Red
            btnConScale3.ForeColor = Color.White
            btnConScale3.Text = "Connect"
            stScale3 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #3 :: IP Address " & _svIPAddress3 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        If My.Computer.Network.Ping(_svIPAddress4) Then
            btnConScale4.BackColor = Color.Green
            btnConScale4.ForeColor = Color.Red
            btnConScale4.Text = "Connected"
            stScale4 = 1
        Else
            btnConScale4.BackColor = Color.Red
            btnConScale4.ForeColor = Color.White
            btnConScale4.Text = "Connect"
            stScale4 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #4 :: IP Address " & _svIPAddress4 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        If My.Computer.Network.Ping(_svIPAddress5) Then
            btnConScale5.BackColor = Color.Green
            btnConScale5.ForeColor = Color.Red
            btnConScale5.Text = "Connected"
            stScale5 = 1
        Else
            btnConScale5.BackColor = Color.Red
            btnConScale5.ForeColor = Color.White
            btnConScale5.Text = "Connect"
            stScale5 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #5 :: IP Address " & _svIPAddress5 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        If My.Computer.Network.Ping(_svIPAddress6) Then
            btnConScale6.BackColor = Color.Green
            btnConScale6.ForeColor = Color.Red
            btnConScale6.Text = "Connected"
            stScale6 = 1
        Else
            btnConScale6.BackColor = Color.Red
            btnConScale6.ForeColor = Color.White
            btnConScale6.Text = "Connect"
            stScale6 = 0
            'MessageBox.Show("โปรแกรมไม่สามารถสื่อสารกับเครื่องชั่ง #6 :: IP Address " & _svIPAddress6 & " ได้ กรุณาตรวจสอบดูอีกครั้ง", "Error Disconnect TCP/IP Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        '--------- Start read modbus TCP/IP --------
        lsvScale.Select()
        tmrConnectIP1.Start()
    End Sub
End Class
