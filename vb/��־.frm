VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form3"
   ScaleHeight     =   5070
   ScaleWidth      =   5460
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin MSComctlLib.ListView LvwReturn 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7011
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Set cnn = Nothing
Set rs = Nothing
Unload Me
End Sub

Private Sub Form_Load()
Dim str12 As String
strSql = "select * from 日志 where 单位名='" & Form1.txtid.text & "'"
 Call connect8(strSql)
Label1.Caption = "共查找" + Trim(rs.RecordCount) + "条记录"
LvwReturn.ListItems.Clear
LvwReturn.ColumnHeaders.Clear
LvwReturn.View = lvwReport
'给控件制作表格标题
LvwReturn.ColumnHeaders.Add , , "单位名", 1300
LvwReturn.ColumnHeaders.Add , , "日期时间", 1200
LvwReturn.ColumnHeaders.Add , , "IP地址", 1200
While Not rs.EOF
    Dim i As Integer
    Dim itmx As ListItem
    Dim intCount As Integer
    Dim text_hj As String
    '如果rs.Fields(0)值不为空
    If Not IsNull(rs.Fields(0)) Then
       Set itmx = LvwReturn.ListItems.Add(, , CStr(rs("单位名")))   '给第一字段单位赋值
    End If
    itmx.Checked = True
    intCount = intCount + 1  'word从第二行开始
    '开始给每行表格进行赋值
    itmx.SubItems(1) = rs("日期和时间") & ""  '从第二格赋值单位级别
    itmx.SubItems(2) = rs("IP地址") & "" ' rs!bz_hj & ""
    rs.MoveNext
 Wend
                
End Sub

