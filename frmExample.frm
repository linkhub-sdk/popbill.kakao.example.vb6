VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� īī���� API SDK VB6 Example"
   ClientHeight    =   11415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17580
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   17580
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame4 
      Caption         =   "�˺� īī���� ���� ���"
      Height          =   8055
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   17295
      Begin VB.TextBox txtResult 
         Height          =   4080
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   59
         Top             =   3480
         Width           =   16815
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "�������� ���"
         Height          =   495
         Left            =   6480
         TabIndex        =   58
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "���ۻ��� Ȯ��"
         Height          =   495
         Left            =   4920
         TabIndex        =   57
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1200
         TabIndex        =   56
         Top             =   3045
         Width           =   3615
      End
      Begin VB.Frame Frame9 
         Caption         =   "īī���� ����"
         Height          =   2895
         Left            =   11280
         TabIndex        =   46
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton btnSearch 
            Caption         =   "���۳��� ��� Ȯ��"
            Height          =   495
            Left            =   2880
            TabIndex        =   54
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_SENDER 
            Caption         =   "�߽Ź�ȣ ���� �˾� URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   53
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "�߽Ź�ȣ ��� Ȯ��"
            Height          =   495
            Left            =   2880
            TabIndex        =   52
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_BOX 
            Caption         =   "���۳��� ��ȸ �˾� URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   51
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnListATSTemplate 
            Caption         =   "�˸��� ���ø� ��� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   50
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_TEMPLATE 
            Caption         =   "�˸��� ���ø� ���� �˾� URL"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnListPlusFriendID 
            Caption         =   "�÷���ģ�� ��� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_PLUSFRIEND 
            Caption         =   "�÷���ģ�� �������� �˾� URL"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "ģ���� �̹��� ����"
         Height          =   855
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   5415
         Begin VB.CommandButton btnSendFMS_multi 
            Caption         =   "���� 1000�� ����"
            Height          =   495
            Left            =   3480
            TabIndex        =   45
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS_same 
            Caption         =   "�뷮 1000�� ����"
            Height          =   495
            Left            =   1680
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS 
            Caption         =   "1�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "ģ���� �ؽ�Ʈ ����"
         Height          =   855
         Left            =   5640
         TabIndex        =   38
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendFTS_one 
            Caption         =   "1�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton btnSendFTS_same 
            Caption         =   "�뷮 1000�� ����"
            Height          =   495
            Left            =   1680
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFTS_multi 
            Caption         =   "���� 1000�� ����"
            Height          =   495
            Left            =   3480
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "�˸��� ����"
         Height          =   855
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendATS_multi 
            Caption         =   "���� 1000�� ����"
            Height          =   495
            Left            =   3480
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_same 
            Caption         =   "�뷮 1000�� ����"
            Height          =   495
            Left            =   1680
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_one 
            Caption         =   "1�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   3075
         TabIndex        =   32
         Top             =   360
         Width           =   3105
      End
      Begin VB.Label Label4 
         Caption         =   "������ȣ : "
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����ð�(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   435
         Width           =   2790
      End
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2295
      TabIndex        =   25
      Text            =   "1234567890"
      Top             =   300
      Width           =   1935
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6240
      TabIndex        =   24
      Text            =   "testkorea"
      Top             =   285
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   17400
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         ClipControls    =   0   'False
         Height          =   1935
         Left            =   11520
         TabIndex        =   22
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "���۴ܰ�"
         Height          =   1935
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   4920
         Begin VB.CommandButton btnGetChargeInfo_FMS 
            Caption         =   "ģ���� �̹��� ��������"
            Height          =   410
            Left            =   2520
            TabIndex        =   30
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_FTS 
            Caption         =   "ģ���� �ؽ�Ʈ ��������"
            Height          =   410
            Left            =   2520
            TabIndex        =   29
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost_FMS 
            Caption         =   "ģ���� �̹��� ���۴ܰ�"
            Height          =   410
            Left            =   150
            TabIndex        =   28
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost_ATS 
            Caption         =   "�˸��� ���۴ܰ�"
            Height          =   410
            Left            =   150
            TabIndex        =   21
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost_FTS 
            Caption         =   "ģ���� �ؽ�Ʈ ���۴ܰ�"
            Height          =   410
            Index           =   0
            Left            =   150
            TabIndex        =   20
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_ATS 
            Caption         =   "�˸��� ��������"
            Height          =   410
            Left            =   2520
            TabIndex        =   19
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   13440
         TabIndex        =   10
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "ȸ������ ����"
         Height          =   1935
         Left            =   15480
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "�������� ����Ʈ"
         Height          =   1935
         Left            =   6960
         TabIndex        =   4
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1935
         Left            =   9000
         TabIndex        =   1
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "��Ʈ�� ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   2175
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4680
      TabIndex        =   26
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' �˺� īī���� API VB 6.0 SDK Example
'
' - VB6 SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/569/
' - ������Ʈ ���� : 2018-03-14
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' - 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
'
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private KakaoService As New PBKakaoService

'=========================================================================
' ���๮�������� ����մϴ�.
' - ������Ҵ� �������۽ð� 10���������� �����մϴ�.
'=========================================================================
Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' �˸���(ATS) ���� ���������� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetChargeInfo_ATS_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
            
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, ATS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (���۴ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ģ���� �̹���(FMS) ���� ���������� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetChargeInfo_FMS_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FMS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (���۴ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ģ���� �ؽ�Ʈ(FTS) ���� ���������� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetChargeInfo_FTS_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FTS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (���۴ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = KakaoService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(��ǥ�ڼ���) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(��ȣ��) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetMessages_Click()
    Dim sentInfo As PBKakaoSentResult
    Dim tmp As String
    Dim info As PBKakaoSentDetail
    Dim btnInfo As PBKakaoButton
    
    Set sentInfo = KakaoService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
     
    If sentInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "================== ���ۿ�û ������� ================== " + vbCrLf
    tmp = tmp + "contentType (īī���� ����) : " + sentInfo.contentType + vbCrLf
    tmp = tmp + "templateCode (���ø� �ڵ�) : " + sentInfo.templateCode + vbCrLf
    tmp = tmp + "plusFriendID (�÷���ģ�� ���̵�) : " + sentInfo.plusFriendID + vbCrLf
    tmp = tmp + "sendNum (�߽Ź�ȣ) : " + sentInfo.sendNum + vbCrLf
    tmp = tmp + "altContent (��ü���� ����) : " + sentInfo.altContent + vbCrLf
    tmp = tmp + "altSendType (��ü���� ����) : " + sentInfo.altSendType + vbCrLf
    tmp = tmp + "reserveDT (�����Ͻ�) : " + sentInfo.reserveDT + vbCrLf
    tmp = tmp + "adsYN (�������� ����) : " + CStr(sentInfo.adsYN) + vbCrLf
    tmp = tmp + "imageURL (ģ���� �̹��� URL) : " + sentInfo.imageURL + vbCrLf
    tmp = tmp + "sendCnt (���۰Ǽ�) : " + sentInfo.sendCnt + vbCrLf
    tmp = tmp + "successCnt (�����Ǽ�) : " + sentInfo.successCnt + vbCrLf
    tmp = tmp + "failCnt (���аǼ�) : " + sentInfo.failCnt + vbCrLf
    tmp = tmp + "altCnt (��ü���� �Ǽ�) : " + sentInfo.altCnt + vbCrLf
    tmp = tmp + "cancelCnt (��ҰǼ�) : " + sentInfo.cancelCnt + vbCrLf + vbCrLf
    
    tmp = tmp + "================== ��ư���� ==================" + vbCrLf
    
    For Each btnInfo In sentInfo.btns
        tmp = tmp + "n (��ư��) : " + btnInfo.n + vbCrLf
        tmp = tmp + "t (��ư����) : " + btnInfo.t + vbCrLf
        tmp = tmp + "u1 (��ư��ũ1) : " + btnInfo.u1 + vbCrLf
        tmp = tmp + "u2 (��ư��ũ2) : " + btnInfo.u2 + vbCrLf + vbCrLf
    Next
    
    
    tmp = tmp + vbCrLf + "================== ���۰������ ==================" + vbCrLf
    tmp = tmp + "state | sendDT | result | resultDT | contentType | receiveNum | receiveName | content | altContentType | altSendDT | altResult | altResultDT" + vbCrLf
            
    For Each info In sentInfo.msgs
        
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + CStr(info.result) + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.contentType + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.content + " | "
        tmp = tmp + CStr(info.altContentType) + " | "
        tmp = tmp + info.altSendDT + " | "
        tmp = tmp + CStr(info.altResult) + " | "
        tmp = tmp + info.altResultDT
        tmp = tmp + vbCrLf
        
    Next
        
    txtResult.Text = tmp
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = KakaoService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� �˾� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = KakaoService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = KakaoService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetSenderNumberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim SenderNumber As PBKakaoSenderNumber
    
    Set SenderNumberList = KakaoService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    For Each SenderNumber In SenderNumberList
        tmp = tmp + "�߽Ź�ȣ(number) : " + SenderNumber.number + vbCrLf
        tmp = tmp + "��ǥ��ȣ ��������(representYN) : " + CStr(SenderNumber.representYN) + vbCrLf
        tmp = tmp + "��ϻ���(state) : " + CStr(SenderNumber.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetURL_BOX_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_PLUSFRIEND_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "PLUSFRIEND")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_SENDER_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "SENDER")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_TEMPLATE_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "TEMPLATE")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
 
    '��ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ
    joinData.CorpNum = "0000000403"
    
    '��ǥ�ڼ���
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�
    joinData.addr = "�ּ�"
    
    '����
    joinData.bizType = "����"
    
    '����
    joinData.bizClass = "����"
    
    '���̵�
    joinData.id = "userid"
    
    '��й�ȣ
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ
    joinData.ContactHP = "010-1234-5678"
    
    '����� ����
    joinData.ContactEmail = "test@test.com"
    
    Set Response = KakaoService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub btnListATSTemplate_Click()
    
    Dim tmp As String
    Dim templateList As Collection
    Dim atsTemplate As PBATSTemplate
    Dim btnInfo As PBKakaoButton
    
    
    Set templateList = KakaoService.ListATSTemplate(txtCorpNum.Text)
    
    If templateList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    For Each atsTemplate In templateList
        tmp = tmp + "===================== ���ø� ���� =====================" + vbCrLf
        tmp = tmp + "���ø� �ڵ�(templateCode) : " + atsTemplate.templateCode + vbCrLf
        tmp = tmp + "���ø� ����(templateName) : " + atsTemplate.templateName + vbCrLf
        tmp = tmp + "���ø� ����(template) : " + atsTemplate.template + vbCrLf
        tmp = tmp + "�÷���ģ�� ���̵�(plusFriendID) : " + atsTemplate.plusFriendID + vbCrLf + vbCrLf
        
        If (atsTemplate.btns Is Nothing) = False Then
            For Each btnInfo In atsTemplate.btns
                tmp = tmp + "��ư��(n) : " + btnInfo.n + vbCrLf
                tmp = tmp + "��ư����(t) : " + btnInfo.t + vbCrLf
                tmp = tmp + "��ư��ũ1(u1) : " + btnInfo.u1 + vbCrLf
                tmp = tmp + "��ư��ũ2(u2) : " + btnInfo.u2 + vbCrLf + vbCrLf
            Next
        End If
        
        tmp = tmp + vbCrLf
        
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = KakaoService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnListPlusFriendID_Click()
    Dim plusFriendList As Collection
    Dim tmp As String
    Dim plusFriendInfo As PBPlusFriend
    
    Set plusFriendList = KakaoService.ListPlusFriendID(txtCorpNum.Text)
    
    If plusFriendList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    For Each plusFriendInfo In plusFriendList
        tmp = tmp + "�÷���ģ�� ���̵�(plusFriendID) : " + plusFriendInfo.plusFriendID + vbCrLf
        tmp = tmp + "�÷���ģ�� �̸�(plusFriendName) : " + plusFriendInfo.plusFriendName + vbCrLf
        tmp = tmp + "����Ͻ�(regDT) : " + plusFriendInfo.regDT + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = "testkorea_20161011"
    
    '��й�ȣ
    joinData.pwd = "test@test.com"
    
    '����ڸ�
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
        
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
            
    Set Response = KakaoService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub btnSearch_Click()
    Dim searchList As PBKakaoSearchResult
    Dim SDate As String
    Dim EDate As String
    Dim ReserveYN As String
    Dim state As New Collection
    Dim Item As New Collection
    Dim SenderYN As Boolean
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim tmp As String
    Dim info As PBKakaoSentDetail
    
    '[�ʼ�] ��������, ��������(yyyyMMdd)
    SDate = "20180101"
    
    '[�ʼ�] ��������, ��������(yyyyMMdd)
    EDate = "20181231"
    
    '���ۻ��°� �迭, 0-���, 1-������, 2-����, 3- ��ü, 4-����, 5-���
    state.Add "0"
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    state.Add "5"
    
    '�˻���� �迭, ATS(�˸���), FTS(ģ���� �ؽ�Ʈ), FMS(ģ���� �̹���)
    Item.Add "ATS"
    Item.Add "FTS"
    Item.Add "FMS"
    
    '���๮�� �˻�����, ����(��ü��ȸ), 1(�������� ��ȸ), 0(������� ��ȸ)
    ReserveYN = ""
    
    '������ȸ����, True(������ȸ), False(��ü��ȸ)
    SenderYN = False
    
    '������ ��ȣ
    Page = 1
    
    '������ ��ϰ���, �ִ� 1000��
    PerPage = 50
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"

    Set searchList = KakaoService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order)
     
    If searchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(searchList.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + searchList.message + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(searchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(searchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(searchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(searchList.pageCount) + vbCrLf + vbCrLf
    
    tmp = tmp + "state | sendDT | result | resultDT | contentType | receiveNum | receiveName | content | altContentType | altSendDT | altResult | altResultDT" + vbCrLf
            
    For Each info In searchList.list
        
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + CStr(info.result) + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.contentType + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.content + " | "
        tmp = tmp + CStr(info.altContentType) + " | "
        tmp = tmp + info.altSendDT + " | "
        tmp = tmp + CStr(info.altResult) + " | "
        tmp = tmp + info.altResultDT
        tmp = tmp + vbCrLf
        
    Next
        
    txtResult.Text = tmp
End Sub

Private Sub btnSendATS_multi_Click()
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim ReceiptNum As String
    Dim templateCode As String
    Dim senderNum As String
    Dim altSendType As String
    Dim i As Integer
    
    ' �˸��� ���ø��ڵ� - ListATStemplate API, GetURL(TEMPLATE) API, �Ǵ� �˺�����Ʈ���� Ȯ��
    templateCode = "018020000001"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 99
        
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)

        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
        
        '�˸��� ����, �ִ� 1000��
        rcvInfo.msg = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�." + CStr(i)
        
        '��ü���� �޽��� ����
        rcvInfo.altmsg = "��ü���� �����Դϴ�." + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    ReceiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, senderNum, "", "", altSendType, txtReserveDT.Text, rcvList)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendATS_one_Click()
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim ReceiptNum As String
    Dim templateCode As String
    Dim senderNum As String
    Dim altSendType As String
    
    '�˸��� ���ø��ڵ� - ListATStemplate API, GetURL(TEMPLATE) API, �Ǵ� �˺�����Ʈ���� Ȯ��
    templateCode = "018020000001"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '���Ź�ȣ
    rcvInfo.rcv = "010000111"
    
    '�����ڸ�
    rcvInfo.rcvnm = "�������̸�"
    
    '�˸��� ����, �ִ� 1000��
    rcvInfo.msg = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�."
    
    '��ü���� �޽��� ����
    rcvInfo.altmsg = "��ü���� �׽�Ʈ�Դϴ�."
    
    rcvList.Add rcvInfo
    
    
    ReceiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, senderNum, "", "", altSendType, txtReserveDT.Text, rcvList)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendATS_same_Click()
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim ReceiptNum As String
    Dim templateCode As String
    Dim senderNum As String
    Dim altSendType As String
    Dim content As String
    Dim altContent As String
    Dim i As Integer
    
    ' �˸��� ���ø��ڵ� - ListATStemplate API, GetURL(TEMPLATE) API, �Ǵ� �˺�����Ʈ���� Ȯ��
    templateCode = "018020000001"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '�˸��� ����, �ִ� 1000��
    content = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�."
    
    '��ü���� ����
    altContent = "��ü���� �׽�Ʈ�Դϴ�."
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    ReceiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, senderNum, content, altContent, altSendType, txtReserveDT.Text, rcvList)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFMS_Click()
    Dim ReceiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim filePath As String
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    Dim imageURL As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '÷���̹��� ���ϰ��
    filePath = CommonDialog1.FileName
    
    '�̹��� ��ũ URL
    imageURL = "http://www.popbill.com"
    
    '�÷���ģ�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '���Ź�ȣ
    rcvInfo.rcv = "010000111"
    
    '�����ڸ�
    rcvInfo.rcvnm = "�������̸�"
    
    'ģ���� ����, �ִ� 1000��
    rcvInfo.msg = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�."
    
    '��ü���� �޽��� ����
    rcvInfo.altmsg = "��ü���� �׽�Ʈ�Դϴ�."
    
    rcvList.Add rcvInfo
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList, filePath, imageURL)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFMS_multi_Click()
    Dim ReceiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim filePath As String
    Dim i As Integer
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    Dim imageURL As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '÷���̹��� ���ϰ��
    filePath = CommonDialog1.FileName
    
    '�̹��� ��ũ URL
    imageURL = "http://www.popbill.com"
    
    '�÷���ģ�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
                 
        'ģ���� ����, �ִ� 400��
        rcvInfo.msg = "ģ���� �����Դϴ�. �����ڿ� ���� �ٸ� ������ �����մϴ�." + CStr(i)
        
        '��ü���� �޽��� ����
        rcvInfo.altmsg = "��ü���� �����Դϴ�. �����ڿ� ���� �ٸ� ������ ������ �� �ֽ��ϴ�." + CStr(i)
        
        rcvList.Add rcvInfo
    
    Next
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList, filePath, imageURL)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFMS_same_Click()
    Dim ReceiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim content As String
    Dim altContent As String
    Dim adsYN As Boolean
    Dim filePath As String
    Dim i As Integer
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    Dim imageURL As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '÷���̹��� ���ϰ��
    filePath = CommonDialog1.FileName
    
    '�̹��� ��ũ URL
    imageURL = "http://www.popbill.com"
    
    '�÷���ģ�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    'ģ���� ����, �ִ� 400��
    content = "ģ���� ���� �����Դϴ�. �ִ� 1000�� ���� �Է��� �� �ֽ��ϴ�. ������ ������ ���� �����ڿ��� �����ϴ� �����Դϴ�."
    
    '��ü���� ����
    altContent = "��ü���� �׽�Ʈ�Դϴ�."
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, content, altContent, altSendType, txtReserveDT.Text, adsYN, rcvList, btnList, filePath, imageURL)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFTS_multi_Click()
    Dim ReceiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim i As Integer
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    
    '�÷���ģ�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
                 
        'ģ���� ����, �ִ� 1000��
        rcvInfo.msg = "ģ���� �����Դϴ�. �����ڿ� ���� �ٸ� ������ �����մϴ�." + CStr(i)
        
        '��ü���� �޽��� ����
        rcvInfo.altmsg = "��ü���� �����Դϴ�. �����ڿ� ���� �ٸ� ������ ������ �� �ֽ��ϴ�." + CStr(i)
        
        rcvList.Add rcvInfo
    
    Next
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFTS_one_Click()
    Dim ReceiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    
    
    '�÷���ģ�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '���Ź�ȣ
    rcvInfo.rcv = "010000111"
    
    '�����ڸ�
    rcvInfo.rcvnm = "�������̸�"
    
    'ģ���� ����, �ִ� 1000��
    rcvInfo.msg = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�."
    
    '��ü���� �޽��� ����
    rcvInfo.altmsg = "��ü���� �׽�Ʈ�Դϴ�."
    
    rcvList.Add rcvInfo
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFTS_same_Click()
    Dim ReceiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim content As String
    Dim altContent As String
    Dim adsYN As Boolean
    Dim i As Integer
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    
    '�÷���ģ�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042993"
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    'ģ���� ����, �ִ� 1000��
    content = "ģ���� ���� �����Դϴ�. �ִ� 1000�� ���� �Է��� �� �ֽ��ϴ�. ������ ������ ���� �����ڿ��� �����ϴ� �����Դϴ�."
    
    '��ü���� ����
    altContent = "��ü���� �׽�Ʈ�Դϴ�."
    
    '��ü���� ��������, ����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, senderNum, content, altContent, altSendType, txtReserveDT.Text, adsYN, rcvList, btnList)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

'=========================================================================
' �˸���(ATS) ���۴ܰ��� Ȯ���մϴ�.
'=========================================================================
Private Sub btnUnitCost_ATS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, ATS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "ATS ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' ģ���� �̹���(FMS) ���۴ܰ��� Ȯ���մϴ�.
'=========================================================================
Private Sub btnUnitCost_FMS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FMS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "FMS ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' ģ���� �ؽ�Ʈ(FTS) ���۴ܰ��� Ȯ���մϴ�.
'=========================================================================
Private Sub btnUnitCost_FTS_Click(Index As Integer)
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FTS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "FTS ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-4304-2991"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '��ü��ȸ����, Ture-ȸ����ȸ, False-������ȸ
    joinData.searchAllAllowYN = True
    
    Set Response = KakaoService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�
    CorpInfo.addr = "����Ư����"
    
    '����
    CorpInfo.bizType = "����"
    
    '����
    CorpInfo.bizClass = "����"
    
    Set Response = KakaoService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub Form_Load()
    KakaoService.Initialize LinkID, SecretKey
    
    '����ȯ�� ������ True-���߿�, False-�����
    KakaoService.IsTest = True
End Sub
