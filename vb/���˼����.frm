VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "����λ��������������"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8985
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7920
      TabIndex        =   16
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7920
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������"
      Height          =   1935
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   2055
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "��"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "�����ڵ�:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "��"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "���Ͻڵ�:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "��"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "�����߽ڵ�:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "����|����)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "�������"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   8400
      Top             =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ���"
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin MSComctlLib.ListView LvwReturn 
      Height          =   4935
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8705
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
   Begin VB.Label Label7 
      Caption         =   "��"
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   1800
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub Command1_Click()
   Dim timeS As Double
   Dim ECHO As ICMP_ECHO_REPLY
   Dim s As String
   Dim str_nz, str_ip As String
   Dim str_1, str_2, str_3 As Integer
   Dim str_5, str_6 As String
   Dim Today
   Dim i As Integer
   Dim itmx As ListItem
   Dim intCount As Integer
   Dim text_hj As String
   str_2 = 0
   str_3 = 0
   LvwReturn.ListItems.Clear
   LvwReturn.ColumnHeaders.Clear
   LvwReturn.View = lvwReport
   LvwReturn.ColumnHeaders.Add , , "����״̬", 1200
   LvwReturn.ColumnHeaders.Add , , "��λ��", 1500
   LvwReturn.ColumnHeaders.Add , , "IP��ַ", 1200
   LvwReturn.ColumnHeaders.Add , , "���ʱ��", 1200
   LvwReturn.ColumnHeaders.Add , , "��Ҫ�̶�", 1200
   strSql = "select * from d_ping"
   Call connect8(strSql)
   str_1 = Val(rs.RecordCount)
   For i = 1 To str_1
   s = Trim(rs![IP��ַ])
   Call Ping(s, ECHO)
   '���rs.Fields(0)ֵ��Ϊ��
   'While Not rs.EOF
   If rs.EOF <> True Then
        If Not IsNull(rs.Fields(0)) Then
               If ECHO.status <> IP_SUCCESS Then
                  Set itmx = LvwReturn.ListItems.Add(, , CStr("��ͨ"))   '����һ�ֶε�λ��ֵ
                  itmx.Checked = True
                  intCount = intCount + 1  'word�ӵڶ��п�ʼ
                  str_2 = str_2 + 1
                  itmx.SubItems(1) = rs("��λ��") & ""  '�ӵڶ���ֵ��λ����
                  itmx.SubItems(2) = rs("IP��ַ") & ""
                  itmx.SubItems(3) = rs("���ʱ��") & ""
                  itmx.SubItems(4) = rs("��Ҫ�̶�") & ""
                  str_5 = Trim(rs![��λ��])
                  str_6 = itmx.SubItems(2)
                  Today = Now()
                  rs.MoveNext
                  '���粻ͨ��Ϊ��¼
                  strSql_1 = "select * from ��־"
                  Call connect9(strSql_1)
                  rs_1.AddNew
                  rs_1![��λ��] = str_5
                  rs_1![IP��ַ] = str_6
                  rs_1![���ں�ʱ��] = Now()
                  rs_1.Update
                  Set cnn_1 = Nothing
                  Set rs_1 = Nothing
             Else
                  Set itmx = LvwReturn.ListItems.Add(, , CStr("����"))   '����һ�ֶε�λ��ֵ
                  itmx.Checked = True
                  intCount = intCount + 1  'word�ӵڶ��п�ʼ
                  str_3 = str_3 + 1
                  itmx.SubItems(1) = rs("��λ��") & ""  '�ӵڶ���ֵ��λ����
                  itmx.SubItems(2) = rs("IP��ַ") & ""
                  itmx.SubItems(3) = rs("���ʱ��") & ""
                  itmx.SubItems(4) = rs("��Ҫ�̶�") & ""
                  rs.MoveNext
        End If
   End If
 End If
 Next
   Text1.Text = str_2 + str_3
   Text2.Text = str_2
   Text3.Text = str_3
End Sub

Private Sub Command3_Click()
Set Conn = Nothing
Set rs = Nothing
Unload Me
End Sub

Private Sub LvwReturn_DblClick()
Dim oldText As String
oldText = myListField
oldText = Trim(LvwReturn.SelectedItem.ListSubItems(1))
txtid.Text = oldText
strSql = "select * from ��־ where ��λ��='" & txtid.Text & "'"
 Call connect8(strSql)
 If rs.RecordCount = 0 Then
        MsgBox "�õ�λû���жϼ�¼", , "���˼��ϵͳ"
        Set cnn = Nothing
        Set rs = Nothing
        Unload Me
 Else
       Set cnn = Nothing
       Set rs = Nothing
       Load Form3
       Form3.Show (vbModal)
       Set cnn = Nothing
       Set rs = Nothing
End If
End Sub
