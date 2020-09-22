VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3780
      Top             =   630
   End
   Begin VB.Line dummy 
      Visible         =   0   'False
      X1              =   1995
      X2              =   420
      Y1              =   2835
      Y2              =   1260
   End
   Begin VB.Line real_line 
      X1              =   1155
      X2              =   2310
      Y1              =   1260
      Y2              =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   420
      Shape           =   3  'Circle
      Top             =   420
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lw As Integer
Dim Dgr As Integer
Dim y1 As Integer
Dim y2 As Integer
Dim x1 As Integer
Dim x2 As Integer


Private Sub Form_Load()
    dummy.y1 = dummy.y2 - dummy.y1
    lw = dummy.x2 - dummy.x1
    x1 = dummy.x1
    x2 = dummy.x2
    y1 = dummy.y1
    y2 = dummy.y2
End Sub


Private Sub Timer1_Timer()
    Dgr = Dgr + 5
    If Dgr > 360 Then Dgr = 1
    dummy.y2 = y2 - ((lw / 2) * Sin(ToRadians(Dgr)))
    dummy.x2 = x2 - ((lw / 2) * (1 - Cos(ToRadians(Dgr))))
    real_line.x2 = dummy.x2
    real_line.y2 = dummy.y2
End Sub


Private Function ToRadians(ByVal Deg As Integer) As Single
    Const P = 3.14159
    ToRadians = Deg * 2 * P / 360
End Function

