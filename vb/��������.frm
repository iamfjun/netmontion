VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form2"
   ScaleHeight     =   2685
   ScaleWidth      =   6570
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "报警设置"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Text            =   "Text7"
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "报警设置"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "中断节点数:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "短信服务器:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "接收短信手机(多手机号';'分隔)"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
