VERSION 5.00
Begin VB.Form SmsObj 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10770
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Left            =   8760
      Top             =   6120
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   9840
      TabIndex        =   26
      Text            =   "Text5"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   7440
      TabIndex        =   24
      Text            =   "Text4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   5400
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3480
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNick 
      Height          =   270
      Left            =   5160
      TabIndex        =   10
      Text            =   "�ű��"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtMobile 
      Height          =   270
      Left            =   5160
      TabIndex        =   8
      Text            =   "15958742738"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "���ж���ʾ����"
      Height          =   3015
      Left            =   3840
      TabIndex        =   5
      Top             =   3000
      Width           =   6495
      Begin VB.CommandButton SmsToRTX 
         Caption         =   "����"
         Height          =   735
         Left            =   5160
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtMsg 
         Height          =   975
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtCUin 
         Height          =   270
         Left            =   1320
         TabIndex        =   11
         Text            =   "zbh"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "�������ݣ�"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "������Uin��"
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�������سƣ�"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�ֻ����룺"
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ʾ���ڣ�"
      Height          =   3375
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   3735
      Begin VB.TextBox txtRecvSmsMsg 
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton StopApp 
      Caption         =   "ֹͣӦ��"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton RegApp 
      Caption         =   "����Ӧ��"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtSvrIP 
      Height          =   270
      Left            =   1800
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label15 
      Caption         =   "�������룺"
      Height          =   255
      Left            =   8760
      TabIndex        =   25
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "��¼���룺"
      Height          =   255
      Left            =   6480
      TabIndex        =   23
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "��¼�û���"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "�ӿڱ��룺"
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "����IP��"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�������سƣ�"
      Height          =   180
      Left            =   4080
      TabIndex        =   12
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RTX AppServer IP��"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   184
      Width           =   1620
   End
End
Attribute VB_Name = "SmsObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const App_Name = "mysms"
Dim RootObj As RTXSAPIRootObj '����һ��������
Dim WithEvents SmsObj As RTXSAPISmsObj '����һ�����Ŷ���
Attribute SmsObj.VB_VarHelpID = -1
Dim UserManagerObj As RTXSAPIUserManager '����һ���û��������

Dim SmsReceivers() As String ' �������Ž���������
Dim nMax As Integer
Dim strCn As String

Dim MT_table, MO_table, RPT_table As String
Dim apicode As String

Dim WithEvents AppObj As RTXSAPIObj '����һ����Ϣ����
Attribute AppObj.VB_VarHelpID = -1


Private Sub SmsObj_OnSendSmsMessage(ByVal szSender As String, ByVal szSmsSender As String, ByVal szRecvMobiles As String, ByVal szMsg As String, ByVal szCookie As String)

    On Error GoTo errHandler
    Dim cn4rtx As New ADODB.Connection
    Set cn4rtx = CreateObject("ADODB.Connection")
    cn4rtx.Open "Driver={SQL Server};Server=10.0.0.121,1433;Uid=sa;pwd=mflogin;Database=rtxdb"
    Dim cn As New ADODB.Connection
    Set cn = CreateObject("ADODB.Connection")
    cn.Open strCn
    
    SmsReceivers = Split(szRecvMobiles, ";") '��Ⱥ��ʱ��������Ϊ����ֻ����룬��ÿ���ֻ��������SmsReceivers����
    nMax = UBound(SmsReceivers) '��ȡ�������Ԫ�ظ���
    For i = 0 To nMax 'ѭ���·����Ÿ�ÿ���ֻ�����
    
        If Len(szMsg) >= 1000 Then
            SmsObj.ReSendSmsMessage szSender, szCookie, "<Result>" & "<Item Mobile=" & """" & SmsReceivers(i) & """" & " bSuccess=" & """" & "1" & """" & " Result=" & """" & "��Ϣ����ʧ�ܣ����������������ޡ�" & """" & "/></Result>"
            Exit For
        End If
    '    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & szSender + szSmsSender + SmsReceivers(i) + szMsg + szCookie & vbCrLf '��ӡ��������
        Dim smID As Long
        smID = 1
        'zbh&&13500000000&zzz:yyyyyyyyyyyyyy&1;0;1005;0;0
        Dim sql As String
        Dim rs As New ADODB.Recordset
        sql = "select isnull(max(id),0) as smID from zbh_smssend"
        Set rs = cn4rtx.Execute(sql)
        If Not rs.EOF Then
            smID = rs("smID") + 1
        End If
        cn4rtx.Execute ("insert into zbh_smssend (id,szSender,szSmsSender,SmsReceivers,szMsg,szCookie,sendTime) values (" & smID & ",'" & szSender & "','" & szSmsSender & "','" & SmsReceivers(i) & "','" & szMsg & "','" & szCookie & "','" & Now() & "')")
        
        sql = "select 1 from SYS_User where Mobile='" & SmsReceivers(i) & "'"
        Set rs = CreateObject("ADODB.recordset")
        rs.Open sql, cn4rtx, 1, 1
        If Not rs.EOF Then
            sql = "insert into " + MT_table + " (SM_ID,SRC_ID,MOBILES,CONTENT,SEND_TIME) values (" & smID & "," & smID & ",'" & SmsReceivers(i) & "','" & szMsg & "','" & Now() & "')"
            cn.Execute (sql)
            SmsObj.ReSendSmsMessage szSender, szCookie, "<Result>" & "<Item Mobile=" & """" & SmsReceivers(i) & """" & " bSuccess=" & """" & "0" & """" & " Result=" & """" & "��Ϣ�Ѿ��ɹ�����" & """" & "/></Result>" '���ط��ͽ�����ͻ��ˣ�40���ڱ��뷵�ؽ��������ͻ�����ʾ��ʱ
        Else
            SmsObj.ReSendSmsMessage szSender, szCookie, "<Result>" & "<Item Mobile=" & """" & SmsReceivers(i) & """" & " bSuccess=" & """" & "1" & """" & " Result=" & """" & "��Ϣ����ʧ�ܣ���Ӧ�ֻ�������ϵ�������С�" & """" & "/></Result>" '���ط��ͽ�����ͻ��ˣ�40���ڱ��뷵�ؽ��������ͻ�����ʾ��ʱ
        End If
        
    Next
    cn.Close
    Set cn = Nothing
    cn4rtx.Close
    Set cn4rtx = Nothing
    Exit Sub
errHandler:
'    cn.Close
    Set cn = Nothing
'    cn4rtx.Close
    Set cn4rtx = Nothing
    MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description

End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then '�жϳ����Ƿ������У����Ϊ��
        Title = App.Title '��¼�±�����ı���
        Call MsgBox(App.Title + "(�汾 " & App.Major & "." & App.Minor & "." & App.Revision & ")�����У�", vbInformation) '����һ����ʾ�Ի���
        App.Title = "�ڶ���ִ��" '�ı䱾�������,�����Ų��ἤ���Լ�
        AppActivate Title '������ǰ��ִ�еĳ���
        End '��������ִ��֮ʵ��,��֤������һ������
    End If

    If UCase(Trim(App.EXEName)) <> UCase(Trim(App_Name)) Then
        MsgBox "���뽫���Ų�������Ƹ���Ϊ: " + App_Name
        End
    End If

'    If FindProcess("mysms.exe") Then '�ڴ��޸�Ϊ��Ҫ�ҵĳ�����
'        MsgBox "�ó����������У�"
'        End
'    End If
    
    '�趨Ĭ��ֵ
    Text1.Text = "10.0.0.7"
    Text2.Text = "115"
    Text3.Text = "rtx"
    Text4.Text = "rtx"
    Text5.Text = "99"
             
    Set RootObj = CreateObject("RTXSAPIRootObj.RTXSAPIRootObj") '����������
    Set SmsObj = RootObj.CreateAPISmsObj 'ͨ�������󴴽����Ŷ���
    Set UserManagerObj = RootObj.UserManager 'ͨ�������󴴽��û��������

    SmsObj.AppAction = AA_COPY ' ���ù��˶���
    SmsObj.AppPriority = 0 '����Ӧ��Ȩ��
    
'������Ϣ����
Set AppObj = RootObj.CreateAPIObj   '����Ӧ�ö���

    
    Call RegApp_Click

End Sub
Private Sub Form_Unload(Cancel As Integer)
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "���ڹرգ���ȴ�������" & vbCrLf
    Call StopApp_Click
End Sub

Private Sub RegApp_Click()

    On Error GoTo errHandler
    
    strCn = "DRIVER={MySQL ODBC 3.51 Driver};" & _
             "SERVER=" & Text1.Text & ";" & _
             "DATABASE=mas;" & _
             "UID=" & Text3.Text & ";PWD=" & Text4.Text & ";" & _
             "OPTION=35;"
             ' �������ݿ�
'    Set cn = CreateObject("ADODB.Connection")
'    cn.Open strCn
    
    MT_table = "api_mt_" & Text2.Text
    MO_table = "api_mo_" & Text2.Text
    RPT_table = "api_rpt_" & Text2.Text
    apicode = Text2.Text
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    
    '���ż�¼
    SmsObj.AppGUID = "{54567dfb-9d81-4382-8f0a-4301761029bf}" '����Ӧ��GUID
    SmsObj.AppName = "Tencent.RTX.Sms" ' ����Ӧ������
    SmsObj.ServerIP = txtSvrIP.Text  ' ���÷�������ַ
    SmsObj.ServerPort = "8006" '���÷������˿�
    SmsObj.SmsWordLimit = "60" ' ����2007��ʽ�������Ľӿڣ��������ÿͻ���ÿ�����ų���ΪXX���ֽ�
    SmsObj.RegisterApp 'ע��Ӧ��
    
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "����ע��ɹ�" & vbCrLf
    
    '��Ϣ��¼
    AppObj.ServerIP = txtSvrIP.Text  '���÷�������ַ
    AppObj.ServerPort = "8006" '���÷������˿�
    AppObj.AppName = "IMzbh" '����Ӧ����
    AppObj.AppGUID = "{9FEF6E5D-136C-4b2c-83A5-25B05FDBAC02}" '����Ӧ��GUID
    AppObj.AppAction = AA_COPY
    AppObj.FilterAppName = "all"  '���ù���Ӧ����
    AppObj.FilterRequestType = "Tencent.RTX.IM" '���ù�����Ϣ����
    AppObj.FilterResponseType = "none" '������Ϣ�ظ�����
    AppObj.FilterSender = "anyone" '������Ϣ������
    AppObj.FilterReceiver = "anyone" ' ������Ϣ������
    AppObj.FilterReceiverState = "anystate" ' ������Ϣ������״̬
    AppObj.FilterKey = "" '���ùؼ��֣���Ϊ��ʱ��ʾ����������Ϣ
    AppObj.RegisterApp
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "��Ϣע��ɹ�" & vbCrLf
    
    SmsObj.StartApp "", 4 '����Ӧ��
    '����5����һ�����ж���
    Timer1.Interval = 3000
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "���������ɹ�" & vbCrLf
    AppObj.StartApp "", 4
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "��Ϣ�����ɹ�" & vbCrLf
    
    Exit Sub
errHandler:
    
    MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description


End Sub

Private Sub SmsToRTX_Click() '���ж���

On Error GoTo errHandler

SmsObj.SendMoSmsMessage txtMobile.Text, txtNick.Text, txtMsg.Text, txtCUin.Text '���ж��ţ��Ѷ��ŷ����ͻ���

Exit Sub

errHandler:
MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description

End Sub

Private Sub StopApp_Click()

On Error GoTo errHandler
SmsObj.StopApp 'ֹͣӦ��
'ֹͣ���м��
Timer1.Interval = 0
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "����ֹͣ�ɹ�" & vbCrLf
AppObj.StopApp
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "��Ϣֹͣ�ɹ�" & vbCrLf
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
SmsObj.UnRegisterApp 'ע��Ӧ��
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "����ע��Ӧ��" & vbCrLf
AppObj.UnRegisterApp 'ע��Ӧ��
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "��Ϣע��Ӧ��" & vbCrLf

Exit Sub

errHandler:
MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description

End Sub

Private Sub Timer1_Timer()
    
    On Error GoTo errHandler
    Dim cn4rtx As New ADODB.Connection
    Set cn4rtx = CreateObject("ADODB.Connection")
    cn4rtx.Open "Driver={SQL Server};Server=10.0.0.121,1433;Uid=sa;pwd=mflogin;Database=rtxdb"
    Dim cn As New ADODB.Connection
    Set cn = CreateObject("ADODB.Connection")
    cn.Open strCn
    '�������п�ʼ
    Dim rs4rcv, rs4mo As New ADODB.Recordset
    Dim sql4rcv As String
    sql4rcv = "select * from " & MO_table
    Set rs4rcv = cn.Execute(sql4rcv)
    If Not rs4rcv.EOF Then
    'ѭ�����б�
        While (Not rs4rcv.EOF)
            Set rs4mo = cn4rtx.Execute("select zbh_smssend.*,SYS_User.ID as CUin from zbh_smssend,SYS_User where szSender=SYS_User.UserName and zbh_smssend.id=" & rs4rcv("sm_id"))
            If Not rs4mo.EOF Then
            '����ж�Ӧ����ԭsmid
                If RootObj.QueryUserState(rs4mo("szSender")) <> "Offline" Then
                '����Ӧ��Ա�����߻��뿪״̬��
                    '���͵��ͻ���
                    SmsObj.SendMoSmsMessage rs4rcv("MOBILE"), rs4mo("szSender"), rs4rcv("CONTENT"), rs4mo("CUin") '���ж��ţ��Ѷ��ŷ����ͻ���
                    '��ӵ�rtx���ݿ�
                    cn4rtx.Execute ("insert into zbh_smsrecieve (ID,SM_ID,MOBILE,[CONTENT],MO_TIME,rtx_mo_time) values (" & rs4rcv("auto_sn") & "," & rs4rcv("sm_id") & ",'" & rs4rcv("MOBILE") & "','" & rs4rcv("CONTENT") & "','" & rs4rcv("MO_TIME") & "','" & Now() & "')")
                    'ɾ��mas���ݿ������б��ж�Ӧ��¼
                    cn.Execute ("delete from " & MO_table & " where auto_sn=" & rs4rcv("auto_sn"))
                End If
            Else
            '���û��ԭsmid
            'ֱ��ɾ�����ж���
            cn.Execute ("delete from " & MO_table & " where auto_sn=" & rs4rcv("auto_sn"))
            End If
            rs4rcv.MoveNext
        Wend
    End If
    Set rs4rcv = Nothing
    cn4rtx.Close
    Set cn4rtx = Nothing
    cn.Close
    Set cn = Nothing
    Exit Sub

errHandler:
    Set rs4rcv = Nothing
    'cn4rtx.Close
    Set cn4rtx = Nothing
'    cn.Close
    Set cn = Nothing
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "Error # " & Str(Err.Number) & Err.Description & vbCrLf
End Sub

Private Sub AppObj_OnRecvMessage(ByVal Message As RTXSAPILib.IRTXSAPIMessage) '���յ���Ϣʱ�������¼�
    On Error GoTo errHandler
    Dim cn4rtx As New ADODB.Connection
    Set cn4rtx = CreateObject("ADODB.Connection")
    cn4rtx.Open "Driver={SQL Server};Server=10.0.0.121,1433;Uid=sa;pwd=mflogin;Database=rtxdb"
    cn4rtx.Execute ("insert into zbh_messagesave (Sender,Receivers,Contents,RSTime) values ('" & Message.Sender & "','" & Message.Receivers & "','" & Replace(Message.Content, "'", "''") & "','" & Now() & "')")
    cn4rtx.Close
    Set cn4rtx = Nothing
    Exit Sub
errHandler:
'    cn4rtx.Close
    Set cn4rtx = Nothing
    MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description
End Sub


Function FindProcess(ProcessName) As Boolean
Dim ps
'ö�ٽ���
For Each ps In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_ 'ѭ������
If UCase(ps.Name) = UCase(ProcessName) Then
    FindProcess = True
    Exit Function
End If
Next
End Function
