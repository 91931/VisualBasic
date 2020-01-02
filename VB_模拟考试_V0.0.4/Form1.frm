VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "2019年全国青少年禁毒知识答题活动题库"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15795
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   15795
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1200
      TabIndex        =   36
      Text            =   "38,39,58,60,138,146,239,251,363,390,437,438,457,471,473,496,535,536,574,"
      Top             =   650
      Width           =   14415
   End
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   9360
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "Form1.frx":048A
      Top             =   7680
      Width           =   6255
   End
   Begin VB.ComboBox Combo_End_No 
      Height          =   300
      Left            =   2400
      TabIndex        =   34
      Top             =   262
      Width           =   855
   End
   Begin VB.ComboBox Combo_Start_No 
      Height          =   300
      Left            =   1200
      TabIndex        =   33
      Top             =   262
      Width           =   855
   End
   Begin VB.TextBox Text_Error 
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Top             =   240
      Width           =   7455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   6720
      TabIndex        =   27
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6495
      Left            =   10680
      TabIndex        =   26
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   11456
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   8040
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"Form1.frx":05A9
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   10455
      Begin VB.Label Label14 
         Caption         =   "题"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "当前第"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label_No 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1440
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   7080
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "已答"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "题"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   4320
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "耗时"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   7560
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6720
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "正确率"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   8520
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "秒"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   9240
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "选项A"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   8535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "选项B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   3030
         Width           =   8535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "选项C"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   3660
         Width           =   8535
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "选项D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   4290
         Width           =   8535
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "选项E"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   4920
         Width           =   7695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   10095
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "结束答题"
      Height          =   975
      Left            =   7320
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   8760
      Top             =   120
   End
   Begin VB.CommandButton Command5 
      Caption         =   "开始答题"
      Height          =   975
      Left            =   5325
      TabIndex        =   4
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "下一题"
      Height          =   975
      Left            =   3315
      TabIndex        =   3
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "上一题"
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "读取题库"
      Height          =   345
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "错题重做"
      Height          =   360
      Left            =   3600
      TabIndex        =   0
      Top             =   232
      Width           =   1395
   End
   Begin VB.Label Label12 
      Caption         =   "-"
      Height          =   180
      Left            =   2160
      TabIndex        =   29
      Top             =   315
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "起始题号"
      Height          =   180
      Left            =   360
      TabIndex        =   25
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r_output As New ADODB.Recordset
Dim v_right_number As Long
Dim v_all_number As Long

Private Sub Command1_Click()
'   CommonDialog1.Filter = "Rich Text Format files|*.*"
'   CommonDialog1.ShowOpen
'   RichTextBox1.LoadFile CommonDialog1.FileName

    Dim v_error_array() As String
    v_error_array() = Split(Me.Text2.Text, ",")

    If Not r_output.BOF And Not r_output.EOF Then
        For i = 1 To r_output.RecordCount
            r_output.MoveFirst
            r_output.Delete
        Next i
    End If

    RichTextBox1.LoadFile "禁毒.txt"
    v_read_number = 578
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = v_read_number
    ProgressBar1.Value = 0
    
        For i = Me.Combo_Start_No.Text To Me.Combo_Start_No.Text + v_read_number - 1
            v_find_id = InStr(i, RichTextBox1.Text, CStr(i) + ".", vbTextCompare)
            v_find_next_id = InStr(i, RichTextBox1.Text, CStr(i + 1) + ".", vbTextCompare)
            v_find_qustion = InStr(v_find_id, RichTextBox1.Text, "A.", vbTextCompare) - 2
            v_find_answer_a = InStr(v_find_qustion, RichTextBox1.Text, " B.", vbTextCompare) - 2
            
            v_find_answer_b = InStr(v_find_answer_a, RichTextBox1.Text, " C.", vbTextCompare) - 2
            If v_find_answer_b > v_find_next_id Then
                v_find_answer_b = InStr(v_find_answer_a, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            End If
            
            
            v_find_answer_c = InStr(v_find_answer_b, RichTextBox1.Text, " D.", vbTextCompare) - 2
            If v_find_answer_c > v_find_next_id Or v_find_answer_c < 0 Then
                v_find_answer_c = InStr(v_find_answer_b, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            End If
            
            
            v_find_answer_d = InStr(v_find_answer_c, RichTextBox1.Text, " E.", vbTextCompare) - 2
            If v_find_answer_d > v_find_next_id Or v_find_answer_d < 0 Then
                v_find_answer_d = InStr(v_find_answer_c, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            End If
            
            v_find_answer_e = InStr(v_find_answer_d, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            
    
            
            If v_find_answer_d > v_find_next_id Then
                v_find_answer_d = 0
            End If
            If v_find_answer_e > v_find_next_id Then
                v_find_answer_e = 0
            End If
            
            If v_find_id > 0 Then
                For j = 1 To UBound(v_error_array())
                    If v_error_array(j - 1) = i Then
                        r_output.AddNew
                        
                        r_output.Fields("ID") = CStr(i)
                        r_output.Fields("题目") = Trim(Mid(RichTextBox1.Text, v_find_id + Len(CStr(i)) + 1, v_find_qustion - v_find_id - 2))
                        r_output.Fields("选项A") = Trim(Mid(RichTextBox1.Text, v_find_qustion + 1, v_find_answer_a - v_find_qustion))
                        r_output.Fields("选项B") = Trim(Mid(RichTextBox1.Text, v_find_answer_a + 1, v_find_answer_b - v_find_answer_a))
                        If v_find_answer_c > 0 Then
                            r_output.Fields("选项C") = Trim(Mid(RichTextBox1.Text, v_find_answer_b + 1, v_find_answer_c - v_find_answer_b))
                        End If
                        If v_find_answer_d > 0 Then
                            r_output.Fields("选项D") = Trim(Mid(RichTextBox1.Text, v_find_answer_c + 1, v_find_answer_d - v_find_answer_c))
                        End If
                        If v_find_answer_e > 0 Then
                            r_output.Fields("选项E") = Trim(Mid(RichTextBox1.Text, v_find_answer_d + 1, v_find_answer_e - v_find_answer_d))
                        End If
                        If v_find_next_id > 0 Then
                            r_output.Fields("正确答案") = Trim(Mid(RichTextBox1.Text, v_find_next_id - 3, 1))
                        Else
                            r_output.Fields("正确答案") = Trim(Right(RichTextBox1.Text, 3))
                        End If
                    End If
                Next j

            End If
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
        Next i
    
    Me.ProgressBar1.Visible = False
    
    r_output.MoveFirst
    Show_qustion
    If Not r_output.BOF And Not r_output.EOF Then
        Me.Command3.Enabled = True
        Me.Command4.Enabled = True
        Me.Command5.Enabled = True
        Me.Command6.Enabled = True
    End If
End Sub


Private Sub Command2_Click()
    If Me.Combo_End_No.Text - Me.Combo_Start_No.Text < 0 Then
        MsgBox "请正确选择出题范围！！！"
        Exit Sub
    End If

    If Not r_output.BOF And Not r_output.EOF Then
        For i = 1 To r_output.RecordCount
            r_output.MoveFirst
            r_output.Delete
        Next i
    End If

    RichTextBox1.LoadFile "禁毒.txt"
    v_read_number = Me.Combo_End_No.Text - Me.Combo_Start_No.Text + 1
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = v_read_number
    ProgressBar1.Value = 0
    
        For i = Me.Combo_Start_No.Text To Me.Combo_Start_No.Text + v_read_number - 1
            v_find_id = InStr(i, RichTextBox1.Text, CStr(i) + ".", vbTextCompare)
            v_find_next_id = InStr(i, RichTextBox1.Text, CStr(i + 1) + ".", vbTextCompare)
            v_find_qustion = InStr(v_find_id, RichTextBox1.Text, "A.", vbTextCompare) - 2
            v_find_answer_a = InStr(v_find_qustion, RichTextBox1.Text, " B.", vbTextCompare) - 2
            
            v_find_answer_b = InStr(v_find_answer_a, RichTextBox1.Text, " C.", vbTextCompare) - 2
            If v_find_answer_b > v_find_next_id Then
                v_find_answer_b = InStr(v_find_answer_a, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            End If
            
            
            v_find_answer_c = InStr(v_find_answer_b, RichTextBox1.Text, " D.", vbTextCompare) - 2
            If v_find_answer_c > v_find_next_id Or v_find_answer_c < 0 Then
                v_find_answer_c = InStr(v_find_answer_b, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            End If
            
            
            v_find_answer_d = InStr(v_find_answer_c, RichTextBox1.Text, " E.", vbTextCompare) - 2
            If v_find_answer_d > v_find_next_id Or v_find_answer_d < 0 Then
                v_find_answer_d = InStr(v_find_answer_c, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            End If
            
            v_find_answer_e = InStr(v_find_answer_d, RichTextBox1.Text, " 答案", vbTextCompare) - 2
            
    
            
            If v_find_answer_d > v_find_next_id Then
                v_find_answer_d = 0
            End If
            If v_find_answer_e > v_find_next_id Then
                v_find_answer_e = 0
            End If
            
            If v_find_id > 0 Then
                r_output.AddNew
                r_output.Fields("ID") = CStr(i)
                r_output.Fields("题目") = Trim(Mid(RichTextBox1.Text, v_find_id + Len(CStr(i)) + 1, v_find_qustion - v_find_id - 2))
                r_output.Fields("选项A") = Trim(Mid(RichTextBox1.Text, v_find_qustion + 1, v_find_answer_a - v_find_qustion))
                r_output.Fields("选项B") = Trim(Mid(RichTextBox1.Text, v_find_answer_a + 1, v_find_answer_b - v_find_answer_a))
                If v_find_answer_c > 0 Then
                    r_output.Fields("选项C") = Trim(Mid(RichTextBox1.Text, v_find_answer_b + 1, v_find_answer_c - v_find_answer_b))
                End If
                If v_find_answer_d > 0 Then
                    r_output.Fields("选项D") = Trim(Mid(RichTextBox1.Text, v_find_answer_c + 1, v_find_answer_d - v_find_answer_c))
                End If
                If v_find_answer_e > 0 Then
                    r_output.Fields("选项E") = Trim(Mid(RichTextBox1.Text, v_find_answer_d + 1, v_find_answer_e - v_find_answer_d))
                End If
                If v_find_next_id > 0 Then
                    r_output.Fields("正确答案") = Trim(Mid(RichTextBox1.Text, v_find_next_id - 3, 1))
                Else
                    r_output.Fields("正确答案") = Trim(Right(RichTextBox1.Text, 3))
                End If
            End If
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
        Next i
    
    Me.ProgressBar1.Visible = False
    
    r_output.MoveFirst
    Show_qustion
    If Not r_output.BOF And Not r_output.EOF Then
        Me.Command3.Enabled = True
        Me.Command4.Enabled = True
        Me.Command5.Enabled = True
        Me.Command6.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
If Not r_output.BOF Then
    r_output.MovePrevious
    Show_qustion
End If
End Sub

Private Sub Command4_Click()
If Not r_output.EOF Then
    r_output.MoveNext
    Show_qustion
    v_all_number = v_all_number + 1
End If
End Sub

Private Sub Show_qustion()
    If Not r_output.BOF And Not r_output.EOF Then
    
        Me.Option1.Value = False
        Me.Option2.Value = False
        Me.Option3.Value = False
        Me.Option4.Value = False
        Me.Option5.Value = False
        Me.Label_No.Caption = ""
        
        Me.Option1.ForeColor = &H80000012
        Me.Option2.ForeColor = &H80000012
        Me.Option3.ForeColor = &H80000012
        Me.Option4.ForeColor = &H80000012
        Me.Option5.ForeColor = &H80000012
        
        If r_output.Fields("选项A") = "" Then Me.Option1.Visible = False Else Me.Option1.Visible = True
        If r_output.Fields("选项B") = "" Then Me.Option2.Visible = False Else Me.Option2.Visible = True
        If r_output.Fields("选项C") = "" Then Me.Option3.Visible = False Else Me.Option3.Visible = True
        If r_output.Fields("选项D") = "" Then Me.Option4.Visible = False Else Me.Option4.Visible = True
        If r_output.Fields("选项E") = "" Then Me.Option5.Visible = False Else Me.Option5.Visible = True
    
        Me.Label_No.Caption = r_output.Fields("ID")
        Me.Text1.Text = Chr(13) + Chr(10) + "    " + r_output.Fields("题目")
        Me.Option1.Caption = r_output.Fields("选项A")
        Me.Option2.Caption = r_output.Fields("选项B")
        Me.Option3.Caption = r_output.Fields("选项C")
        Me.Option4.Caption = r_output.Fields("选项D")
        Me.Option5.Caption = r_output.Fields("选项E")
    End If
End Sub

Private Sub Command5_Click()
    v_all_number = 0
    v_right_number = 0
    
    Me.Label1.Caption = 0
    Me.Label3.Caption = 0
    Me.Label3.Caption = 0
    Me.Timer1.Interval = 1000
    r_output.MoveFirst
    Show_qustion
End Sub

Private Sub Command6_Click()
    Me.Timer1.Interval = 0
End Sub

Private Sub DataGrid1_Click()
    Show_qustion
End Sub

Private Sub Form_Load()
    r_output.CursorLocation = adUseClient
    r_output.Fields.Append "ID", adInteger
    r_output.Fields.Append "题目", adVarChar, 500
    r_output.Fields.Append "选项A", adVarChar, 200
    r_output.Fields.Append "选项B", adVarChar, 200
    r_output.Fields.Append "选项C", adVarChar, 200
    r_output.Fields.Append "选项D", adVarChar, 200
    r_output.Fields.Append "选项E", adVarChar, 200
    r_output.Fields.Append "正确答案", adVarChar, 200
    r_output.Fields.Append "我的选择", adVarChar, 200

    r_output.Open

    Set Me.DataGrid1.DataSource = r_output
    v_all_number = 0
    v_right_number = 0
    
    Me.Option1.Visible = False
    Me.Option2.Visible = False
    Me.Option3.Visible = False
    Me.Option4.Visible = False
    Me.Option5.Visible = False
    
    Me.ProgressBar1.Visible = False
    
    Me.Combo_Start_No.AddItem "1"
    Me.Combo_Start_No.AddItem "101"
    Me.Combo_Start_No.AddItem "201"
    Me.Combo_Start_No.AddItem "301"
    Me.Combo_Start_No.AddItem "401"
    Me.Combo_Start_No.AddItem "501"
    Me.Combo_Start_No.SelText = "1"
    
    Me.Combo_End_No.AddItem "100"
    Me.Combo_End_No.AddItem "200"
    Me.Combo_End_No.AddItem "300"
    Me.Combo_End_No.AddItem "400"
    Me.Combo_End_No.AddItem "500"
    Me.Combo_End_No.AddItem "578"
    Me.Combo_End_No.SelText = "100"
    
    Me.Command3.Enabled = False
    Me.Command4.Enabled = False
    Me.Command5.Enabled = False
    Me.Command6.Enabled = False
    
End Sub

Private Sub Option1_Click()
    If Not r_output.BOF And Not r_output.EOF Then
        r_output.Fields("我的选择") = "A"
        If r_output.Fields("正确答案") = "A" Then
            Me.Option1.ForeColor = &HFF0000
            v_right_number = v_right_number + 1
        Else
            Me.Option1.ForeColor = &HFF&
            Me.Text_Error.Text = Me.Text_Error.Text + CStr(r_output.Fields("ID")) + ","
            If r_output.Fields("正确答案") = "A" Then Me.Option1.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "B" Then Me.Option2.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "C" Then Me.Option3.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "D" Then Me.Option4.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "E" Then Me.Option5.ForeColor = &HFF0000
        End If
    End If
End Sub

Private Sub Option2_Click()
    If Not r_output.BOF And Not r_output.EOF Then
        r_output.Fields("我的选择") = "B"
        If r_output.Fields("正确答案") = "B" Then
            Me.Option2.ForeColor = &HFF0000
            v_right_number = v_right_number + 1
        Else
            Me.Option2.ForeColor = &HFF&
            Me.Text_Error.Text = Me.Text_Error.Text + CStr(r_output.Fields("ID")) + ","
            If r_output.Fields("正确答案") = "A" Then Me.Option1.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "B" Then Me.Option2.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "C" Then Me.Option3.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "D" Then Me.Option4.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "E" Then Me.Option5.ForeColor = &HFF0000
        End If
    End If
End Sub

Private Sub Option3_Click()
    If Not r_output.BOF And Not r_output.EOF Then
        r_output.Fields("我的选择") = "C"
        If r_output.Fields("正确答案") = "C" Then
            Me.Option3.ForeColor = &HFF0000
            v_right_number = v_right_number + 1
        Else
            Me.Option3.ForeColor = &HFF&
            Me.Text_Error.Text = Me.Text_Error.Text + CStr(r_output.Fields("ID")) + ","
            If r_output.Fields("正确答案") = "A" Then Me.Option1.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "B" Then Me.Option2.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "C" Then Me.Option3.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "D" Then Me.Option4.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "E" Then Me.Option5.ForeColor = &HFF0000
        End If
    End If
End Sub

Private Sub Option4_Click()
    If Not r_output.BOF And Not r_output.EOF Then
        r_output.Fields("我的选择") = "D"
        If r_output.Fields("正确答案") = "D" Then
            Me.Option4.ForeColor = &HFF0000
            v_right_number = v_right_number + 1
        Else
            Me.Option4.ForeColor = &HFF&
            Me.Text_Error.Text = Me.Text_Error.Text + CStr(r_output.Fields("ID")) + ","
            If r_output.Fields("正确答案") = "A" Then Me.Option1.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "B" Then Me.Option2.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "C" Then Me.Option3.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "D" Then Me.Option4.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "E" Then Me.Option5.ForeColor = &HFF0000
        End If
    End If
End Sub

Private Sub Option5_Click()
    If Not r_output.BOF And Not r_output.EOF Then
        r_output.Fields("我的选择") = "E"
        If r_output.Fields("正确答案") = "E" Then
            Me.Option5.ForeColor = &HFF0000
            v_right_number = v_right_number + 1
        Else
            Me.Option5.ForeColor = &HFF&
            Me.Text_Error.Text = Me.Text_Error.Text + CStr(r_output.Fields("ID")) + ","
            If r_output.Fields("正确答案") = "A" Then Me.Option1.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "B" Then Me.Option2.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "C" Then Me.Option3.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "D" Then Me.Option4.ForeColor = &HFF0000
            If r_output.Fields("正确答案") = "E" Then Me.Option5.ForeColor = &HFF0000
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Me.Label1.Caption = Me.Label1.Caption + 1
    Me.Label9.Caption = v_all_number
    Me.Label3.Caption = v_right_number
    Me.Label10.Caption = v_all_number
End Sub
