VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConnector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connector Line"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlArrow 
      Left            =   360
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":0000
            Key             =   "TOP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":03CE
            Key             =   "LEFT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":0798
            Key             =   "RIGHT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":0B64
            Key             =   "BOTTOM"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the Line to Select"
      Height          =   195
      Left            =   2370
      TabIndex        =   2
      Top             =   645
      Width           =   1860
   End
   Begin VB.Image imgArrow 
      Height          =   240
      Index           =   1
      Left            =   2205
      Picture         =   "frmConnector.frx":0F33
      Top             =   1725
      Width           =   240
   End
   Begin VB.Line lneRight 
      BorderWidth     =   2
      Index           =   1
      X1              =   2880
      X2              =   4245
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line lneMiddle 
      BorderWidth     =   2
      Index           =   1
      X1              =   2910
      X2              =   2910
      Y1              =   1860
      Y2              =   3030
   End
   Begin VB.Line lneLeft 
      BorderWidth     =   2
      Index           =   1
      X1              =   1620
      X2              =   2895
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Label lblObject2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Here And Move The Object"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   675
      Left            =   4230
      TabIndex        =   1
      Top             =   2730
      Width           =   1155
   End
   Begin VB.Label lblObject1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Here And Move The Object"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   675
      Left            =   465
      TabIndex        =   0
      Top             =   1470
      Width           =   1155
   End
End
Attribute VB_Name = "frmConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Developed by: V.Senthamil Selvan
'''Date: 15-Oct-2002
'''''''''''''''''''''''''''
'''Description:
'''     The form has array of label, Image and Line controls. For the demo purpose
'''First instance of the controls are coded. Programmer can instantiate
'''any number of labels and lines at runtime. Passing any source and target
'''label to the moveConnectLine method will draw the connect line between them
'''The arrow direction is based on the passed source and target label to
'''moveConnectLine method. To draw connectline with respect to the current X,Y
'''cursor movement call drawConnectLine method with X,Y pos on the form
'''Mouse move event. While dynamically creating Line please use same index
'''for imgArrow, Line and label.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim bStartMove As Boolean   ''' move a window wherever clicked
Dim sXClickPos As Single    ''' Current x pos
Dim sYClickPos As Single    ''' current Y Pos
Dim objLine As clsLine

Private Sub Form_Load()
Set objLine = New clsLine
'imlArrow.MaskColor = vbWhite
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lineSelected As Line
Set lineSelected = objLine.selectLine(Me, X, Y)
If Not lineSelected Is Nothing Then
    lneLeft(lineSelected.Index).BorderWidth = 3
    lneRight(lineSelected.Index).BorderWidth = 3
    lneMiddle(lineSelected.Index).BorderWidth = 3
Else
    lneLeft(1).BorderWidth = 2
    lneRight(1).BorderWidth = 2
    lneMiddle(1).BorderWidth = 2
End If
End Sub

Private Sub lblObject1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    bStartMove = True
    ''' get the mouse click X,Y position
    sXClickPos = X
    sYClickPos = Y
Else
    bStartMove = False
End If
End Sub

Private Sub lblObject1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bStartMove Then
    '''''' Move the object with respect to the clicked position
    If sXClickPos > X Then
        lblObject1.Left = lblObject1.Left - (sXClickPos - X)
   ElseIf sXClickPos < X Then
        lblObject1.Left = lblObject1.Left + (X - sXClickPos)
   End If
   If sYClickPos < Y Then
        lblObject1.Top = lblObject1.Top + (Y - sYClickPos)
   ElseIf sYClickPos > Y Then
        lblObject1.Top = lblObject1.Top - (sYClickPos - Y)
   End If
   '''''''''
   objLine.moveConnectLine lblObject1, lblObject2, 1, Me
End If
End Sub

Private Sub lblObject1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bStartMove = False
End Sub

Private Sub lblObject2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    bStartMove = True
    ''' get the mouse click X,Y position
    sXClickPos = X
    sYClickPos = Y
Else
    bStartMove = False
End If
End Sub

Private Sub lblObject2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bStartMove Then
    If sXClickPos > X Then
        lblObject2.Left = lblObject2.Left - (sXClickPos - X)
   ElseIf sXClickPos < X Then
        lblObject2.Left = lblObject2.Left + (X - sXClickPos)
   End If
   If sYClickPos < Y Then
        lblObject2.Top = lblObject2.Top + (Y - sYClickPos)
   ElseIf sYClickPos > Y Then
        lblObject2.Top = lblObject2.Top - (sYClickPos - Y)
   End If
   objLine.moveConnectLine lblObject2, lblObject1, 1, Me
End If
End Sub

Private Sub lblObject2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bStartMove = False
End Sub
