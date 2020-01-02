VERSION 5.00
Begin VB.Form FileRename_Form_PreviewImg 
   Caption         =   "新建文件夹"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   1935
   Icon            =   "FileRename_Form_PreviewImg.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   1935
   StartUpPosition =   2  '屏幕中心
   Begin VB.Image Image_Next 
      Height          =   480
      Left            =   1320
      Picture         =   "FileRename_Form_PreviewImg.frx":048A
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Image_Pre 
      Height          =   480
      Left            =   0
      Picture         =   "FileRename_Form_PreviewImg.frx":09D4
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "FileRename_Form_PreviewImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
    If UCase(Right(r_output.Fields("原始文件名称"), 4)) = ".JPG" Then

        Me.Image1.Picture = LoadPicture(r_output.Fields("原始文件名称"), vbLPLarge, vbLPColor)
        
        Me.Height = Screen.Height - 900
        Me.Width = Me.Height * (Me.Image1.Picture.Width / Me.Image1.Picture.Height)
    
        
        Me.Image1.Width = Me.Width
        Me.Image1.Height = Me.Height
        Me.Image_Next.Left = Me.Width - 1000
        Me.Caption = r_output.Fields("原始文件名称")
        Me.Left = (Screen.Width - Me.Width) / 2
        
        Me.Image_Pre.Top = (Me.Height - Me.Image_Pre.Height) / 2
        Me.Image_Next.Top = (Me.Height - Me.Image_Next.Height) / 2

    End If
End Sub



Private Sub Image_Next_Click()
    If Not r_output.EOF And Not r_output.BOF Then
        r_output.MoveNext
        If r_output.EOF Then
            r_output.MoveFirst
        End If

        Me.Image1.Picture = LoadPicture(r_output.Fields("原始文件名称"), vbLPLarge, vbLPColor)
        Me.Width = Me.Height * (Me.Image1.Picture.Width / Me.Image1.Picture.Height)

        Me.Image1.Width = Me.Width
        Me.Image1.Height = Me.Height
        Me.Image_Next.Left = Me.Width - 1000
        Me.Caption = r_output.Fields("原始文件名称")
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
End Sub

Private Sub Image_Pre_Click()
    If Not r_output.EOF And Not r_output.BOF Then
        r_output.MovePrevious
        If r_output.BOF Then
            r_output.MoveLast
        End If

        Me.Image1.Picture = LoadPicture(r_output.Fields("原始文件名称"), vbLPLarge, vbLPColor)
        Me.Width = Me.Height * (Me.Image1.Picture.Width / Me.Image1.Picture.Height)

        Me.Image1.Width = Me.Width
        Me.Image1.Height = Me.Height
        Me.Image_Next.Left = Me.Width - 1000
        Me.Caption = r_output.Fields("原始文件名称")
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
End Sub
