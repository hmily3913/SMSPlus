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
   StartUpPosition =   3  '窗口缺省
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
      Text            =   "张斌和"
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
      Caption         =   "上行短信示例："
      Height          =   3015
      Left            =   3840
      TabIndex        =   5
      Top             =   3000
      Width           =   6495
      Begin VB.CommandButton SmsToRTX 
         Caption         =   "发送"
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
         Caption         =   "短信内容："
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "接收者Uin："
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "接收者呢称："
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "手机号码："
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "提示窗口："
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
      Caption         =   "停止应用"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton RegApp 
      Caption         =   "启动应用"
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
      Caption         =   "短信子码："
      Height          =   255
      Left            =   8760
      TabIndex        =   25
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "登录密码："
      Height          =   255
      Left            =   6480
      TabIndex        =   23
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "登录用户："
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "接口编码："
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "短信IP："
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "接收者呢称："
      Height          =   180
      Left            =   4080
      TabIndex        =   12
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RTX AppServer IP："
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
Dim RootObj As RTXSAPIRootObj '声明一个根对象
Dim WithEvents SmsObj As RTXSAPISmsObj '声明一个短信对象
Attribute SmsObj.VB_VarHelpID = -1
Dim UserManagerObj As RTXSAPIUserManager '声明一个用户管理对象

Dim SmsReceivers() As String ' 声明短信接收者数组
Dim nMax As Integer
Dim strCn As String

Dim MT_table, MO_table, RPT_table As String
Dim apicode As String

Dim WithEvents AppObj As RTXSAPIObj '声明一个消息对象
Attribute AppObj.VB_VarHelpID = -1


Private Sub SmsObj_OnSendSmsMessage(ByVal szSender As String, ByVal szSmsSender As String, ByVal szRecvMobiles As String, ByVal szMsg As String, ByVal szCookie As String)

    On Error GoTo errHandler
    Dim cn4rtx As New ADODB.Connection
    Set cn4rtx = CreateObject("ADODB.Connection")
    cn4rtx.Open "Driver={SQL Server};Server=10.0.0.121,1433;Uid=sa;pwd=mflogin;Database=rtxdb"
    Dim cn As New ADODB.Connection
    Set cn = CreateObject("ADODB.Connection")
    cn.Open strCn
    
    SmsReceivers = Split(szRecvMobiles, ";") '当群发时，接收者为多个手机号码，把每个手机号码放入SmsReceivers数组
    nMax = UBound(SmsReceivers) '获取该数组的元素个数
    For i = 0 To nMax '循环下发短信给每个手机号码
    
        If Len(szMsg) >= 1000 Then
            SmsObj.ReSendSmsMessage szSender, szCookie, "<Result>" & "<Item Mobile=" & """" & SmsReceivers(i) & """" & " bSuccess=" & """" & "1" & """" & " Result=" & """" & "信息发送失败，字数超过允许上限。" & """" & "/></Result>"
            Exit For
        End If
    '    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & szSender + szSmsSender + SmsReceivers(i) + szMsg + szCookie & vbCrLf '打印下行内容
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
            SmsObj.ReSendSmsMessage szSender, szCookie, "<Result>" & "<Item Mobile=" & """" & SmsReceivers(i) & """" & " bSuccess=" & """" & "0" & """" & " Result=" & """" & "信息已经成功发送" & """" & "/></Result>" '返回发送结果给客户端，40秒内必须返回结果，否则客户端提示超时
        Else
            SmsObj.ReSendSmsMessage szSender, szCookie, "<Result>" & "<Item Mobile=" & """" & SmsReceivers(i) & """" & " bSuccess=" & """" & "1" & """" & " Result=" & """" & "信息发送失败，对应手机不在联系人名单中。" & """" & "/></Result>" '返回发送结果给客户端，40秒内必须返回结果，否则客户端提示超时
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
    If App.PrevInstance = True Then '判断程序是否已运行，如果为真
        Title = App.Title '记录下本程序的标题
        Call MsgBox(App.Title + "(版本 " & App.Major & "." & App.Minor & "." & App.Revision & ")已运行！", vbInformation) '产生一个提示对话框
        App.Title = "第二次执行" '改变本程序标题,这样才不会激活自己
        AppActivate Title '激活先前已执行的程序
        End '结束本次执行之实例,保证仅运行一个程序
    End If

    If UCase(Trim(App.EXEName)) <> UCase(Trim(App_Name)) Then
        MsgBox "必须将短信插件的名称更改为: " + App_Name
        End
    End If

'    If FindProcess("mysms.exe") Then '在此修改为你要找的程序名
'        MsgBox "该程序正在运行！"
'        End
'    End If
    
    '设定默认值
    Text1.Text = "10.0.0.7"
    Text2.Text = "115"
    Text3.Text = "rtx"
    Text4.Text = "rtx"
    Text5.Text = "99"
             
    Set RootObj = CreateObject("RTXSAPIRootObj.RTXSAPIRootObj") '创建根对象
    Set SmsObj = RootObj.CreateAPISmsObj '通过根对象创建短信对象
    Set UserManagerObj = RootObj.UserManager '通过根对象创建用户管理对象

    SmsObj.AppAction = AA_COPY ' 设置过滤动作
    SmsObj.AppPriority = 0 '设置应用权限
    
'增加消息对象
Set AppObj = RootObj.CreateAPIObj   '创建应用对象

    
    Call RegApp_Click

End Sub
Private Sub Form_Unload(Cancel As Integer)
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "正在关闭，请等待。。。" & vbCrLf
    Call StopApp_Click
End Sub

Private Sub RegApp_Click()

    On Error GoTo errHandler
    
    strCn = "DRIVER={MySQL ODBC 3.51 Driver};" & _
             "SERVER=" & Text1.Text & ";" & _
             "DATABASE=mas;" & _
             "UID=" & Text3.Text & ";PWD=" & Text4.Text & ";" & _
             "OPTION=35;"
             ' 连接数据库
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
    
    '短信记录
    SmsObj.AppGUID = "{54567dfb-9d81-4382-8f0a-4301761029bf}" '设置应用GUID
    SmsObj.AppName = "Tencent.RTX.Sms" ' 设置应用名称
    SmsObj.ServerIP = txtSvrIP.Text  ' 设置服务器地址
    SmsObj.ServerPort = "8006" '设置服务器端口
    SmsObj.SmsWordLimit = "60" ' 这是2007正式版新增的接口，可以设置客户端每条短信长度为XX个字节
    SmsObj.RegisterApp '注册应用
    
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "短信注册成功" & vbCrLf
    
    '消息记录
    AppObj.ServerIP = txtSvrIP.Text  '设置服务器地址
    AppObj.ServerPort = "8006" '设置服务器端口
    AppObj.AppName = "IMzbh" '设置应用名
    AppObj.AppGUID = "{9FEF6E5D-136C-4b2c-83A5-25B05FDBAC02}" '设置应用GUID
    AppObj.AppAction = AA_COPY
    AppObj.FilterAppName = "all"  '设置过滤应用名
    AppObj.FilterRequestType = "Tencent.RTX.IM" '设置过滤消息类型
    AppObj.FilterResponseType = "none" '设置消息回复类型
    AppObj.FilterSender = "anyone" '设置消息发送者
    AppObj.FilterReceiver = "anyone" ' 设置消息接收者
    AppObj.FilterReceiverState = "anystate" ' 设置消息接收者状态
    AppObj.FilterKey = "" '设置关键字，当为空时表示过滤所有消息
    AppObj.RegisterApp
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "消息注册成功" & vbCrLf
    
    SmsObj.StartApp "", 4 '启动应用
    '设置5秒检查一次上行短信
    Timer1.Interval = 3000
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "短信启动成功" & vbCrLf
    AppObj.StartApp "", 4
    txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "消息启动成功" & vbCrLf
    
    Exit Sub
errHandler:
    
    MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description


End Sub

Private Sub SmsToRTX_Click() '上行短信

On Error GoTo errHandler

SmsObj.SendMoSmsMessage txtMobile.Text, txtNick.Text, txtMsg.Text, txtCUin.Text '上行短信，把短信发给客户端

Exit Sub

errHandler:
MsgBox "Error # " & Str(Err.Number) & Chr(13) & Err.Description

End Sub

Private Sub StopApp_Click()

On Error GoTo errHandler
SmsObj.StopApp '停止应用
'停止上行检测
Timer1.Interval = 0
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "短信停止成功" & vbCrLf
AppObj.StopApp
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "消息停止成功" & vbCrLf
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
SmsObj.UnRegisterApp '注销应用
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "短信注销应用" & vbCrLf
AppObj.UnRegisterApp '注销应用
txtRecvSmsMsg.Text = txtRecvSmsMsg.Text & "消息注销应用" & vbCrLf

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
    '短信上行开始
    Dim rs4rcv, rs4mo As New ADODB.Recordset
    Dim sql4rcv As String
    sql4rcv = "select * from " & MO_table
    Set rs4rcv = cn.Execute(sql4rcv)
    If Not rs4rcv.EOF Then
    '循环上行表
        While (Not rs4rcv.EOF)
            Set rs4mo = cn4rtx.Execute("select zbh_smssend.*,SYS_User.ID as CUin from zbh_smssend,SYS_User where szSender=SYS_User.UserName and zbh_smssend.id=" & rs4rcv("sm_id"))
            If Not rs4mo.EOF Then
            '如果有对应发送原smid
                If RootObj.QueryUserState(rs4mo("szSender")) <> "Offline" Then
                '检查对应人员是在线或离开状态。
                    '发送到客户端
                    SmsObj.SendMoSmsMessage rs4rcv("MOBILE"), rs4mo("szSender"), rs4rcv("CONTENT"), rs4mo("CUin") '上行短信，把短信发给客户端
                    '添加到rtx数据库
                    cn4rtx.Execute ("insert into zbh_smsrecieve (ID,SM_ID,MOBILE,[CONTENT],MO_TIME,rtx_mo_time) values (" & rs4rcv("auto_sn") & "," & rs4rcv("sm_id") & ",'" & rs4rcv("MOBILE") & "','" & rs4rcv("CONTENT") & "','" & rs4rcv("MO_TIME") & "','" & Now() & "')")
                    '删除mas数据库中上行表中对应记录
                    cn.Execute ("delete from " & MO_table & " where auto_sn=" & rs4rcv("auto_sn"))
                End If
            Else
            '如果没有原smid
            '直接删除上行队列
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

Private Sub AppObj_OnRecvMessage(ByVal Message As RTXSAPILib.IRTXSAPIMessage) '接收到消息时触发该事件
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
'枚举进程
For Each ps In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_ '循环进程
If UCase(ps.Name) = UCase(ProcessName) Then
    FindProcess = True
    Exit Function
End If
Next
End Function
