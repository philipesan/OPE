VERSION 5.00
Begin VB.Form frmSessionControl 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lbRh 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbAdmin 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbSessao 
      BackColor       =   &H8000000D&
      Caption         =   "Nenhum Usuário"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmSessionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
