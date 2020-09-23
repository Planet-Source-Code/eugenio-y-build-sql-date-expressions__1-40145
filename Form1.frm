VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDBType 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3960
      List            =   "Form1.frx":0013
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3120
      Width           =   8775
   End
   Begin VB.ComboBox cmbFormat 
      Height          =   315
      ItemData        =   "Form1.frx":0043
      Left            =   3960
      List            =   "Form1.frx":0050
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdBuildString 
      Caption         =   "Build SQL String"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Database type :"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Input format :"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a date :"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuildString_Click()
Dim myDate As Date
Dim objSQLDateExpression As New clsSQLDateExpression
Dim strSql As String
Dim strExpr As String

If Not objSQLDateExpression.ValidateDate(txtDate, "/", cmbFormat.Text, myDate) Then
    MsgBox "Invalid date"
    txtString.Text = ""
    Exit Sub
End If

strExpr = objSQLDateExpression.DateToSQLExpression(myDate, cmbDBType.Text)

strSql = "INSERT INTO TABLE1(dateField1) values(" & strExpr & ")"
strSql = strSql & vbCr & vbLf
strSql = strSql & "SELECT * FROM TABLE1 WHERE dateField1 = " & strExpr
txtString.Text = strSql
End Sub

