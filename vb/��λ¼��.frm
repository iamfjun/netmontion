VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_add 
   Caption         =   "����ڵ���Ϣ"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6375
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "�༭"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin MSComctlLib.ListView LvwReturn 
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5953
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
   Begin VB.Label Label4 
      Caption         =   "��Ҫ�̶�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "���ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "IP��ַ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "��λ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frm_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pIndex As Integer

Private Sub Command1_Click()
If Command5.Caption = "�༭" Then
   '�����¼�¼
    strSql = "select * from d_ping"
    Call connect8(strSql)
    rs.AddNew
    rs![��λ��] = Trim(text(1).text)
    rs![IP��ַ] = Trim(text(0).text)
    rs![���ʱ��] = Trim(text(2).text)
    rs![��Ҫ�̶�] = Trim(text(3).text)
    rs.Update
    MsgBox "���ݱ���ɹ�"
    Call Form_Load
    Command2.Enabled = True
    Command5.Enabled = True
    Command4.Enabled = True
    Command1.Enabled = True
 Else
    strSql = "update * from d_ping"
    strSql = "update d_ping set ��λ��='" & Trim(text(1).text) & "',IP��ַ='" & Trim(text(0).text) & "',���ʱ��='" & Trim(text(2).text) & "',��Ҫ�̶�='" & Trim(text(3).text) & "' where ��λ����='" & Trim(text(1).text) & "' "
    Call connect8(strSql)
    MsgBox "�����޸ĳɹ�"
 End If
End Sub

Private Sub Command2_Click()
Set Conn = Nothing
 Set rs = Nothing
Unload Me
End Sub

Private Sub Command3_Click()
 Dim i As Integer
 For i = 0 To 3
        text(i).text = ""
 Next
 Command2.Enabled = False
 Command5.Enabled = False
 Command4.Enabled = False
 Command1.Enabled = True
 'text(0).text = Now()
End Sub

Private Sub Command4_Click()
Dim str1 As String
Dim str2 As String
Dim rstemp As ADODB.Recordset
str1 = Trim(text(1).text)
ret = MsgBox("Ҫ����������¼��", vbYesNo, "���񳵹������ϵͳ")
   If ret = 6 Then
      strSQL1 = "delete from d_ping where ��λ��='" & str1 & "'"
      Set rstemp = connect8(strSQL1)
      MsgBox "����ɾ���ɹ�!", vbInformation, "���񳵹������ϵͳ"
   Else
      frm_add.Refresh
   End If
End Sub

Private Sub Command5_Click()
Dim i As Integer
 If Command5.Caption = "�༭" Then
    Command5.Caption = "����"
    For i = 0 To 3
           text(i).Locked = False
    Next
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command1.Enabled = True
 Else
     For i = 0 To 3
         text(i).Locked = True
     Next
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command1.Enabled = True
    Command5.Caption = "�༭"
 End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 3
         text(i).Locked = False
     Next
strSql = "select * from d_ping"
Call connect8(strSql)
LvwReturn.ListItems.Clear
LvwReturn.ColumnHeaders.Clear
LvwReturn.View = lvwReport
LvwReturn.ColumnHeaders.Add , , "��Ҫ�̶�", 1200
LvwReturn.ColumnHeaders.Add , , "��λ��", 1500
LvwReturn.ColumnHeaders.Add , , "IP��ַ", 1200
LvwReturn.ColumnHeaders.Add , , "���ʱ��(����)", 1800
While Not rs.EOF
      Dim itmx As ListItem
      Dim intCount As Integer
      Dim text_hj As String
 '���rs.Fields(0)ֵ��Ϊ��
        If Not IsNull(rs.Fields(0)) Then
           Set itmx = LvwReturn.ListItems.Add(, , CStr(rs("��Ҫ�̶�")))   '����һ�ֶε�λ��ֵ
        End If
        itmx.Checked = True
        intCount = intCount + 1  'word�ӵڶ��п�ʼ
        '��ʼ��ÿ�б����и�ֵ
        itmx.SubItems(1) = rs("��λ��") & ""  '�ӵڶ���ֵ��λ����
        itmx.SubItems(2) = rs("IP��ַ") & ""
        itmx.SubItems(3) = rs("���ʱ��") & ""
         rs.MoveNext
  Wend
        Set cnn = Nothing
        Set rs = Nothing
Set cnn = Nothing
Set rs = Nothing
End Sub

Private Sub LvwReturn_ItemClick(ByVal Item As MSComctlLib.ListItem)
  pIndex = Item.Index
End Sub

Private Sub LvwReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If LvwReturn.ListItems.Count = 0 Then Exit Sub
'''If LvwReturn.SelectedItem Then
'''    i = LvwReturn.SelectedItem.Index
'''End If
'If i = 0 Then i = 1
If LvwReturn.ListItems.Item(pIndex).Selected Then
    If LvwReturn.ListItems.Item(pIndex).SubItems(1) <> "" Then
        text(0).text = LvwReturn.ListItems.Item(pIndex).SubItems(2)
        text(1).text = LvwReturn.ListItems.Item(pIndex).SubItems(1)
        text(2).text = LvwReturn.ListItems.Item(pIndex).SubItems(3)
        text(3).text = LvwReturn.ListItems.Item(pIndex).text '"" ' LvwReturn.ListItems.Item(pIndex).SubItems(4)
    End If
End If

End Sub
