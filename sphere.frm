VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create by Fabiana S. Palacios"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Texture Sphere"
      Height          =   555
      Left            =   2460
      TabIndex        =   2
      Top             =   6300
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4560
      TabIndex        =   1
      Top             =   6300
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Empty Sphere"
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   6300
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cX, cY, cR, o

Private Sub Command1_Click()
  Call drawEmptySphere
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Command3_Click()
  Call drawTextureSphere
End Sub

Private Sub Form_Load()
  Call cValue
End Sub

Public Sub cValue()
  cX = Form1.Width / 2
  cY = Form1.Height / 2
    If cX < cY Then
      cR = cY / 2
    Else
      cR = cX / 2
    End If
End Sub

Public Sub drawEmptySphere()
  
  Form1.Refresh
  Form1.DrawStyle = 0
  Form1.FillStyle = 1
  
  o = 1.1
  
  Form1.Line (cX - cR, cY)-(cX + cR, cY)
  For i = 0.1 To 1 Step 0.1
    Form1.Circle (cX, cY), cR, , , , i
  Next i
   
  Form1.Line (cX, cY - cR)-(cX, cY + cR)
  For i = 1 To 3 Step o
     Form1.Circle (cX, cY), cR, , , , o
     o = o * 1.3
     Next i
   For i = 1 To 6 Step o
     Form1.Circle (cX, cY), cR, , , , o
     o = o * 1.8
   Next i
End Sub

Public Sub drawTextureSphere()
Dim b%, r%
b = 255
r = 255
Form1.Refresh
Form1.DrawStyle = 5
Form1.FillStyle = 0

  For i = 1 To 0.1 Step -0.01
    b = b - 2
    r = r - 2
    Form1.Circle (cX, cY), cR, , , , i
    Form1.FillColor = RGB(r, 0, b)
  Next i
   
End Sub
