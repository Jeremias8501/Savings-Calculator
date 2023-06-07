VERSION 5.00
Begin VB.Form frmSavings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Savings Account"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   735
      Index           =   0
      Left            =   1560
      TabIndex        =   8
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtFinal 
      Height          =   855
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtMonths 
      Height          =   855
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtInterest 
      Height          =   855
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtDeposit 
      Height          =   855
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Final Balance"
      Height          =   855
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Number of Months"
      Height          =   855
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Yearly Interest"
      Height          =   855
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Monthly Deposit"
      Height          =   855
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmSavings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Deposit As Single
Dim Interest As Single
Dim Months As Single
Dim Final As Single


Private Sub cmdCalculate_Click(Index As Integer)
Dim IntRate As Single
'Read values from text boxes
Deposit = Val(txtDeposit.Text)
Interest = Val(txtInterest.Text)
IntRate = Interest / 1200
Months = Val(txtMonths.Text)
'Compute final value and put in text box
Final = Deposit * ((1 + IntRate) ^ Months - 1) / IntRate
txtFinal.Text = Format(Final, "#####0.00")
End Sub

Private Sub cmdExit_Click(Index As Integer)
End
End Sub
