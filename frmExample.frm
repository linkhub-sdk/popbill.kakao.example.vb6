VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� īī���� API SDK VB6 Example"
   ClientHeight    =   13350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   ScaleHeight     =   13350
   ScaleWidth      =   18645
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   13320
      TabIndex        =   71
      Top             =   285
      Width           =   4455
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2295
      TabIndex        =   24
      Text            =   "1234567890"
      Top             =   300
      Width           =   1935
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6240
      TabIndex        =   23
      Text            =   "testkorea"
      Top             =   285
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   17760
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         ClipControls    =   0   'False
         Height          =   2415
         Left            =   11880
         TabIndex        =   21
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "���۴ܰ�"
         Height          =   2415
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   4920
         Begin VB.CommandButton btnGetChargeInfo_FMS 
            Caption         =   "ģ���� �̹��� ��������"
            Height          =   410
            Left            =   2520
            TabIndex        =   29
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_FTS 
            Caption         =   "ģ���� �ؽ�Ʈ ��������"
            Height          =   410
            Left            =   2520
            TabIndex        =   28
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetUnitCost_FMS 
            Caption         =   "ģ���� �̹��� ���۴ܰ�"
            Height          =   410
            Left            =   150
            TabIndex        =   27
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetUnitCost_ATS 
            Caption         =   "�˸��� ���۴ܰ�"
            Height          =   410
            Left            =   150
            TabIndex        =   20
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetUnitCost_FTS 
            Caption         =   "ģ���� �ؽ�Ʈ ���۴ܰ�"
            Height          =   410
            Left            =   150
            TabIndex        =   19
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_ATS 
            Caption         =   "�˸��� ��������"
            Height          =   410
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "����� ����"
         Height          =   2415
         Left            =   13800
         TabIndex        =   9
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "����� ���� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "ȸ������ ����"
         Height          =   2415
         Left            =   15840
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "�������� ����Ʈ"
         Height          =   2415
         Left            =   6960
         TabIndex        =   4
         Top             =   240
         Width           =   2295
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   2415
         Left            =   9360
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
   Begin VB.Frame Frame4 
      Caption         =   "�˺� īī���� ���� ���"
      Height          =   9015
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   17775
      Begin VB.Frame Frame10 
         Caption         =   "��û��ȣ �Ҵ� ���۰� ó��"
         Height          =   1335
         Left            =   6600
         TabIndex        =   61
         Top             =   3960
         Width           =   6255
         Begin VB.TextBox txtRequestNum 
            Height          =   315
            Left            =   2400
            TabIndex        =   65
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton btnGetMessagesRN 
            Caption         =   "���ۻ��� Ȯ��"
            Height          =   495
            Left            =   360
            TabIndex        =   63
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton btnCancelReserveRN 
            Caption         =   "�������� ���"
            Height          =   495
            Left            =   3120
            TabIndex        =   62
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "��û��ȣ(requestNum) :"
            Height          =   375
            Left            =   180
            TabIndex        =   64
            Top             =   320
            Width           =   2175
         End
      End
      Begin VB.TextBox txtResult 
         Height          =   3240
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   57
         Top             =   5400
         Width           =   17295
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "�������� ���"
         Height          =   495
         Left            =   3120
         TabIndex        =   56
         Top             =   4560
         Width           =   2775
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "���ۻ��� Ȯ��"
         Height          =   495
         Left            =   360
         TabIndex        =   55
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   2400
         TabIndex        =   54
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Frame Frame9 
         Caption         =   "īī���� ����"
         Height          =   3615
         Left            =   11760
         TabIndex        =   45
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton btnCheckSenderNumber 
            Caption         =   "�߽Ź�ȣ ��� ���� Ȯ��"
            Height          =   495
            Left            =   2880
            TabIndex        =   72
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton btnGetATSTemplate 
            Caption         =   "�˸��� ���ø� ���� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnSearch 
            Caption         =   "���۳��� ��� Ȯ��"
            Height          =   495
            Left            =   2880
            TabIndex        =   53
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSenderNumberMgtURL 
            Caption         =   "�߽Ź�ȣ ���� �˾� URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   52
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "�߽Ź�ȣ ��� Ȯ��"
            Height          =   495
            Left            =   2880
            TabIndex        =   51
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSentListURL 
            Caption         =   "���۳��� ��ȸ �˾� URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   50
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnListATSTemplate 
            Caption         =   "�˸��� ���ø� ��� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CommandButton btnGetATSTemplateMgtURL 
            Caption         =   "�˸��� ���ø� ���� �˾� URL"
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnListPlusFriendID 
            Caption         =   "īī���� ä�� ��� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetPlusFriendMgtURL 
            Caption         =   "īī���� ä�� ���� �˾� URL"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "ģ���� �̹��� ����"
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   5415
         Begin VB.CommandButton btnSendFMS_multi 
            Caption         =   "�뷮 1000�� ����"
            Height          =   495
            Left            =   3480
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS_same 
            Caption         =   "���� 1000�� ����"
            Height          =   495
            Left            =   1680
            TabIndex        =   43
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS_ONE 
            Caption         =   "1�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "ģ���� �ؽ�Ʈ ����"
         Height          =   855
         Left            =   5640
         TabIndex        =   37
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendFTS_one 
            Caption         =   "1�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton btnSendFTS_same 
            Caption         =   "���� 1000�� ����"
            Height          =   495
            Left            =   1680
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFTS_multi 
            Caption         =   "�뷮 1000�� ����"
            Height          =   495
            Left            =   3480
            TabIndex        =   38
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "�˸��� ����"
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendATS_multi 
            Caption         =   "�뷮 1000�� ����"
            Height          =   495
            Left            =   3480
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_same 
            Caption         =   "���� 1000�� ����"
            Height          =   495
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_one 
            Caption         =   "1�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtSndDT 
         Height          =   315
         Left            =   3075
         TabIndex        =   31
         Top             =   360
         Width           =   3105
      End
      Begin VB.Frame Frame14 
         Caption         =   "������ȣ ���� ��� (��û��ȣ ���Ҵ�)"
         Height          =   1335
         Left            =   120
         TabIndex        =   58
         Top             =   3960
         Width           =   6255
         Begin VB.Label Label5 
            Caption         =   "������ȣ(receiptNum) :"
            Height          =   375
            Left            =   180
            TabIndex        =   59
            Top             =   320
            Width           =   2175
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����ð�(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   435
         Width           =   2790
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   12600
      TabIndex        =   70
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
      Height          =   180
      Left            =   240
      TabIndex        =   26
      Top             =   360
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4680
      TabIndex        =   25
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
' �˺� īī���� API VB SDK Example
'
' - ������Ʈ ���� : 2022-04-06
' - ���� ������� ����ó : 1600-9854
' - ���� ������� �̸��� : code@linkhubcorp.com
' - VB SDK ������ �ȳ� : https://docs.popbill.com/kakao/tutorial/vb
'
' <�׽�Ʈ �������� �غ����>
' 1) 30, 33�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ��Ʈ�� ��û�� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
'
' 2) �˸���/ģ������ �����ϱ� ���� �߽Ź�ȣ ��������� �մϴ�. (��Ϲ���� ����Ʈ/API �ΰ��� ����� �ֽ��ϴ�.)
'   - �˺� ����Ʈ �α��� > [����/�ѽ�] > [īī����] > [�߽Ź�ȣ �������] �޴����� ���
'   - getSenderNumberMgtURL API�� ���� ��ȯ�� URL�� �̿��Ͽ� �߽Ź�ȣ ���
'
' 3) �˸���/ģ������ �����ϱ� ���� īī���� ä�θ� ��� �մϴ�. (��Ϲ���� ����Ʈ/API �ΰ��� ����� �ֽ��ϴ�.)
'   - �˺� ����Ʈ �α��� > [����/�ѽ�] > [īī����] > [īī���� ����]  > īī���� ä�� �������� �޴����� ���
'   - GetPlusFriendMgtURL API�� ���� ��ȯ�� URL�� �̿��Ͽ� īī���� ä�� �������� ���
'
' 4) �˸��� ������ �ϱ� ����  �˸��� ���ø��� ��û �մϴ�.  (��Ϲ���� ����Ʈ/API �ΰ��� ����� �ֽ��ϴ�.)
'   - �˺� ����Ʈ �α��� > [����/�ѽ�] > [īī����] > [īī���� ����]  > �˸��� ���ø� ���� �޴����� ���
'   - GetATSTemplateMgtURL API�� ���� ��ȯ�� URL�� �̿��Ͽ� �˸��� ���ø� ���
'=========================================================================

Option Explicit

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'īī���� ���� Ŭ���� ����
Private KakaoService As New PBKakaoService

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#CheckID
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
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� ����
    joinData.id = "userid"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.corpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.bizType = "����"
    
    '����, �ִ� 100��
    joinData.bizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    Set Response = KakaoService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' īī����(�˸���) ���۽� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_ATS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, ATS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�˸���(ATS) ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' īī����(ģ���� �ؽ�Ʈ) ���۽� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_FTS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FTS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "ģ���� �ؽ�Ʈ(FTS) ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' īī����(ģ���� �̹���) ���۽� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_FMS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FMS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "ģ���� �̹���(FMS) ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺� īī����(�˸���) API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_ATS_Click()
    Dim ChargeInfo As PBChargeInfo

    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, ATS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "���۴ܰ� (unitCost) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "�������� (chargeMethod) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "�������� (rateSystem) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� īī����(ģ���� �ؽ�Ʈ) API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_FTS_Click()
    Dim ChargeInfo As PBChargeInfo
        
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FTS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "���۴ܰ� (unitCost) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "�������� (chargeMethod) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "�������� (rateSystem) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� īī����(ģ���� �̹���) API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_FMS_Click()
    Dim ChargeInfo As PBChargeInfo
        
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FMS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "���۴ܰ� (unitCost) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "�������� (chargeMethod) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "�������� (rateSystem) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ (balance) : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()

    Dim URL As String
    
    URL = KakaoService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = KakaoService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = KakaoService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ (balance) : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
    
    URL = KakaoService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
    
    'Internet Explorer Browser ȣ��
    Dim IE As Object
   
    Dim strResult As String
    Dim strSiteName As String
   
    Set IE = CreateObject("InternetExplorer.Application")
    strSiteName = URL
    IE.Navigate strSiteName
    With IE
        .Visible = True     '������â Ȱ��ȭ
        .Resizable = True   '������â ũ�� ���� On/Off
        .MenuBar = True     '�޴��� On/Off
        .Toolbar = True     '���� On/Off
        .AddressBar = True  '�ּҹ� On/Off
        .StatusBar = False  '���¹� On/Off
    End With
   
    Set IE = Nothing
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()

    Dim URL As String
    
    URL = KakaoService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
        
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� ����
    joinData.id = "testkorea"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf#$%123"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
        
    Set Response = KakaoService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub


'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = "testkorea"
    
    Set info = KakaoService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.tel + " | " _
            + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#ListContact
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

    tmp = "id(���̵�) | personName(����) | email(�̸���) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/kakao/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
                
    Set Response = KakaoService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = KakaoService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
            
    tmp = tmp + "��ǥ�ڼ��� (ceoname) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "��ȣ�� (corpName) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "�ּ� (addr) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "���� (bizType) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "���� (bizClass) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�.
' - https://docs.popbill.com/kakao/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.bizClass = "����"
    
    Set Response = KakaoService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
'���ε� ���ø��� ������ �ۼ��Ͽ� 1���� �˸��� ������ �˺��� �����մϴ�.
' - ������ ���ε� ���ø��� ����� �˸��� ���۳���(content)�� �ٸ� ��� ���۽��� ó���˴ϴ�.
' - ���۽��� �� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ� �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendATS
'=========================================================================
Private Sub btnSendATS_ONE_Click()
    Dim templateCode As String
    Dim snd As String
    Dim content As String
    Dim altSendType As String
    Dim receiptNum As String
    Dim requestNum As String
    
    '�˸��� ���ø��ڵ� - ListATStemplate API, GetPlusFriendMgtURL API, �Ǵ� �˺�����Ʈ���� Ȯ��
    templateCode = "022040000005"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    snd = "07043042991"
    
    '�˸��� ����, �ִ� 1000��
    content = "[ �˺� ]" + vbCrLf
    content = content + "��û�Ͻ� #{���ø��ڵ�}�� ���� �ɻ簡 �Ϸ�Ǿ� ���� ó���Ǿ����ϴ�." + vbCrLf
    content = content + "�ش� ���ø����� ���� �����մϴ�." + vbCrLf + vbCrLf
    content = content + "���ǻ��� �����ø� ��Ʈ�ʼ��ͷ� ���ϰ� �����ֽñ� �ٶ��ϴ�." + vbCrLf + vbCrLf
    content = content + "�˺� ��Ʈ�ʼ��� : 1600-8536" + vbCrLf
    content = content + "support@linkhub.co.kr"
    
    
    '��ü���� �������� (����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����)
    altSendType = "A"
    
    'īī���� ��������
    Dim Messages As New Collection
    Dim info As New PBKakaoReceiver
    
    info.msg = content '�˸��� ����, �ִ� 1000��
    info.altsjt = "�˸��� ��ü ���� ����"  '��ü���� ����, ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
    info.altmsg = "�ܰ� �˸��� ��ü ���� �ܰ�"  '��ü���� ����, �ִ� 2000byte
    info.rcv = "010111222"            '���Ź�ȣ
    info.rcvnm = "popbill"            '�����ڸ�
    info.interOPRefKey = ""  '��Ʈ�� ����Ű (������ ���п�)
    
    '�����ڸ��� �ٸ������� ��ư�߰��� �Ʒ��ڵ� ����
    'Set info.buttonList = New Collection
    'Dim detailButton As New PBKakaoButton
    'detailButton.n = "button"
    'detailButton.t = "WL"
    'detailButton.u1 = "test.popbill.com"
    'detailButton.u2 = "www.popbill.com"
    'info.buttonList.Add detailButton
     
    Messages.Add info
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    
    '�˸��� ��ư������ ���ø� ��û�� ������ ��ư������ �����ϰ� �����ϴ� ��� Buttons�� �� �迭�� ó��.
    Dim Buttons As New Collection
    
    '�˸��� ��ư URL�� #{���ø�����}�� �����Ѱ�� ���ø����� ���� �����Ͽ� ��ư���� ����
    'Dim btn As PBKakaoButton
    
    'Set btn = New PBKakaoButton
    
    'btn.n = "��ư��"                        '��ư��
    'btn.t = "WL"                            '��ư���� WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    'btn.u1 = "https://www.linkhub.co.kr"    '�۸�ũ-iOS, ����ũ-Mobile
    'btn.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    'Buttons.Add btn
    
    
    receiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, snd, "", "", altSendType, txtSndDT.Text, Messages, txtUserID.Text, requestNum, Buttons, "")
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' ���ε� ���ø� ������ �ۼ��Ͽ� �ټ����� �˸��� ������ �˺��� �����ϸ�, ��� �����ڿ��� ���� ������ �����մϴ�. (�ִ� 1,000��)
' - ������ ���ε� ���ø��� ����� �˸��� ���۳���(content)�� �ٸ� ��� ���۽��� ó���˴ϴ�.
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendATS
'=========================================================================
Private Sub btnSendATS_SAME_Click()
    Dim templateCode As String
    Dim snd As String
    Dim content As String
    Dim altSubject As String
    Dim altContent As String
    Dim altSendType As String
    Dim receiptNum As String
    Dim i As Integer
    Dim requestNum As String
    
    '�˸��� ���ø��ڵ� - ListATStemplate API, GetPlusFriendMgtURL API, �Ǵ� �˺�����Ʈ���� Ȯ��
    templateCode = "022040000005"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    snd = "07043042991"
    
    '(����) �˸��� ����, �ִ� 1000��
    content = "[ �˺� ]" + vbCrLf
    content = content + "��û�Ͻ� #{���ø��ڵ�}�� ���� �ɻ簡 �Ϸ�Ǿ� ���� ó���Ǿ����ϴ�." + vbCrLf
    content = content + "�ش� ���ø����� ���� �����մϴ�." + vbCrLf + vbCrLf
    content = content + "���ǻ��� �����ø� ��Ʈ�ʼ��ͷ� ���ϰ� �����ֽñ� �ٶ��ϴ�." + vbCrLf + vbCrLf
    content = content + "�˺� ��Ʈ�ʼ��� : 1600-8536" + vbCrLf
    content = content + "support@linkhub.co.kr"
    
    '(����) ��ü���� ����
    ' ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
    altSubject = "�˸��� ��ü���� ����"
    
    '(����) ��ü���� ����, �ִ� 2000byte
    altContent = "�˸��� ��ü ����"
    
    '��ü���� �������� (����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����)
    altSendType = "A"
    
    'īī���� �������� �迭, �ִ� 1000��
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver
    
    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)  '���Ź�ȣ
        info.rcvnm = "popbill_" + CStr(i) '�����ڸ�
        Messages.Add info
    Next
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˸��� ��ư������ ���ø� ��û�� ������ ��ư������ �����ϰ� �����ϴ� ��� Buttons�� �� �迭�� ó��.
    Dim Buttons As New Collection
    
    '�˸��� ��ư URL�� #{���ø�����}�� �����Ѱ�� ���ø����� ���� �����Ͽ� ��ư���� ����
    'Dim btn As PBKakaoButton
    
    'Set btn = New PBKakaoButton
    
    'btn.n = "��ư��"                        '��ư��
    'btn.t = "WL"                            '��ư���� WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    'btn.u1 = "https://www.linkhub.co.kr"    '�۸�ũ-iOS, ����ũ-Mobile
    'btn.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    'Buttons.Add btn
    
    receiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, snd, content, altContent, altSendType, txtSndDT.Text, Messages, txtUserID.Text, requestNum, Buttons, altSubject)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum

End Sub

'=========================================================================
' ���ε� ���ø��� ������ �ۼ��Ͽ� �ټ����� �˸��� ������ �˺��� �����ϸ�, ������ ���� ���� ������ �����մϴ�. (�ִ� 1,000��)
' - ������ ���ε� ���ø��� ����� �˸��� ���۳���(content)�� �ٸ� ��� ���۽��� ó���˴ϴ�.
' - ���۽��� �� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendATS
'=========================================================================
Private Sub btnSendATS_MULTI_Click()
    Dim templateCode As String
    Dim snd As String
    Dim altSendType As String
    Dim receiptNum As String
    Dim i As Integer
    Dim content As String
    Dim requestNum As String
    
    '�˸��� ���ø��ڵ� - ListATStemplate API, GetPlusFriendMgtURL API, �Ǵ� �˺�����Ʈ���� Ȯ��
    templateCode = "022040000005"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    snd = "07043042991"
    
    '(����) �˸��� ����, �ִ� 1000��
    content = "[ �˺� ]" + vbCrLf
    content = content + "��û�Ͻ� #{���ø��ڵ�}�� ���� �ɻ簡 �Ϸ�Ǿ� ���� ó���Ǿ����ϴ�." + vbCrLf
    content = content + "�ش� ���ø����� ���� �����մϴ�." + vbCrLf + vbCrLf
    content = content + "���ǻ��� �����ø� ��Ʈ�ʼ��ͷ� ���ϰ� �����ֽñ� �ٶ��ϴ�." + vbCrLf + vbCrLf
    content = content + "�˺� ��Ʈ�ʼ��� : 1600-8536" + vbCrLf
    content = content + "support@linkhub.co.kr"
    
    '��ü���� �������� (����-������, C-�˸��峻�� ����, A-��ü���ڳ��� ����)
    altSendType = "A"
    
    'īī���� �������� �迭, �ִ� 1000��
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver

    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)                    '���Ź�ȣ
        info.rcvnm = "popbill_" + CStr(i)                   '�����ڸ�
        info.msg = content                                  '�˸��� ����, �ִ� 1000��
        info.altsjt = "�˸��� ��ü ���� ����"               '��ü���� ����, ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
        info.altmsg = "�˸��� �뷮 ��ü �����Դϴ�." + CStr(i)   '��ü���� �޽��� ����, �ִ� 2000byte
        Messages.Add info
    Next
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˸��� ��ư������ ���ø� ��û�� ������ ��ư������ �����ϰ� �����ϴ� ��� Buttons�� �� �迭�� ó��.
    Dim Buttons As New Collection
    
    '�˸��� ��ư URL�� #{���ø�����}�� �����Ѱ�� ���ø����� ���� �����Ͽ� ��ư���� ����
    'Dim btn As PBKakaoButton
    
    'Set btn = New PBKakaoButton
    
    'btn.n = "��ư��"                        '��ư��
    'btn.t = "WL"                            '��ư���� WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    'btn.u1 = "https://www.linkhub.co.kr"    '�۸�ũ-iOS, ����ũ-Mobile
    'btn.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
   
    'Buttons.Add btn
    
    receiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, snd, "", "", altSendType, txtSndDT.Text, Messages, txtUserID.Text, requestNum, Buttons, "")
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' �ؽ�Ʈ�� ������ 1���� ģ���� ������ �˺��� �����մϴ�.
' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendFTS
'=========================================================================
Private Sub btnSendFTS_ONE_Click()
    Dim receiptNum As String
    Dim plusFriendID As String
    Dim snd As String
    Dim content As String
    Dim altContent As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim requestNum As String
    
    'īī���� ä�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    snd = "07043042991"
    
    '��ü���� �������� (����-������ / C-ģ���峻�� ���� / A-��ü���ڳ��� ����)
    altSendType = "C"

    '�������� ����
    adsYN = False
    
    'īī���� �޼��� ����
    Dim Messages As New Collection
    Dim info As New PBKakaoReceiver
    
    info.msg = "ģ���� �ؽ�Ʈ �Դϴ�"            'ģ���� ����, �ִ� 1000��
    info.altsjt = "ģ���� �ؽ�Ʈ ��ü ���� ����" '��ü���� ����, ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
    info.altmsg = "ģ���� �ؽ�Ʈ ��ü ����"      '��ü���� ����, �ִ� 2000byte
    info.rcv = "010000111"                       '���Ź�ȣ
    info.rcvnm = "�������̸�"                    '�����ڸ�
    
    Messages.Add info
        
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Dim Buttons As New Collection
    Dim btn As PBKakaoButton
    
    Set btn = New PBKakaoButton
    
    btn.n = "��ư��"                        '��ư��
    btn.t = "WL"                            '��ư���� WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    btn.u1 = "http://www.linkhub.co.kr"     '�۸�ũ-iOS, ����ũ-Mobile
    btn.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    Buttons.Add btn
    
    Set btn = New PBKakaoButton
    
    btn.n = "�޽�������"
    btn.t = "MD"
    
    Buttons.Add btn
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, snd, "", "", altSendType, txtSndDT.Text, adsYN, Messages, Buttons, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' �ؽ�Ʈ�� ������ �ټ����� ģ���� ������ �˺��� �����ϸ�, ��� �����ڿ��� ���� ������ �����մϴ�. (�ִ� 1,000��)
' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendFTS
'=========================================================================
Private Sub btnSendFTS_SAME_Click()
    Dim receiptNum As String
    Dim plusFriendID As String
    Dim snd As String
    Dim content As String
    Dim altSubject As String
    Dim altContent As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim i As Integer
    Dim requestNum As String
    
    'īī���� ä�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    snd = "07043042991"
    
    '(����) ģ���� ����, �ִ� 1000��
    content = "ģ���� �ؽ�Ʈ �Դϴ�"
    
    '(����) ��ü���� ����
    ' ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
    altSubject = "ģ���� ��ü���� ����"
    
    '(����) ��ü���� ����, �ִ� 2000byte
    altContent = "ģ���� �ؽ�Ʈ ��ü ����"
        
    '��ü���� �������� (����-������ / C-ģ���峻�� ���� / A-��ü���ڳ��� ����)
    altSendType = "C"

    '�������� ����
    adsYN = False
    
    '�������� �迭, �ִ� 1000��
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver
    
    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)    '���Ź�ȣ
        info.rcvnm = "popbill_" + CStr(i)   '�����ڸ�
        Messages.Add info
    Next
        
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Dim Buttons As New Collection
    Dim btn As PBKakaoButton
    
    Set btn = New PBKakaoButton
    
    btn.n = "��ư��"                        '��ư��
    btn.t = "WL"                            '��ư���� WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    btn.u1 = "http://www.linkhub.co.kr"     '�۸�ũ-iOS, ����ũ-Mobile
    btn.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    Buttons.Add btn
    
    Set btn = New PBKakaoButton
    
    btn.n = "�޽�������"
    btn.t = "MD"
    
    Buttons.Add btn
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, snd, content, altContent, altSendType, txtSndDT.Text, adsYN, Messages, Buttons, txtUserID.Text, requestNum, altSubject)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' �ؽ�Ʈ�� ������ �ټ����� ģ���� ������ �˺��� �����ϸ�, ������ ���� ���� ������ �����մϴ�. (�ִ� 1,000��)
' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendFTS
'=========================================================================
Private Sub btnSendFTS_MULTI_Click()
    Dim receiptNum As String
    Dim plusFriendID As String
    Dim snd As String
    Dim altSendType As String
    Dim adsYN As Boolean
    Dim requestNum As String
    Dim i As Integer
    
    'īī���� ä�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    snd = "07043042991"
    
    '��ü���� �������� (����-������ / C-ģ���峻�� ���� / A-��ü���ڳ��� ����)
    altSendType = "A"

    '�������� ����
    adsYN = False
    
    '�������� �迭, �ִ� 1000��
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver
    
    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)                   '���Ź�ȣ
        info.rcvnm = "popbill_" + CStr(i)                  '�����ڸ�
        info.msg = "�׽�Ʈ ���ø� �Դϴ�"                  '�˸��� ����, �ִ� 1000��
        info.altsjt = "ģ���� ��ü ���� ����"              '��ü���� ����, ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
        info.altmsg = "ģ���� ��ü �����Դϴ�." + CStr(i)  '��ü���� �޽��� ����, �ִ� 2000byte
        Messages.Add info
    Next
            
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Dim Buttons As New Collection
    Dim btn As PBKakaoButton
    
    Set btn = New PBKakaoButton
    
    btn.n = "��ư��"                        '��ư��
    btn.t = "WL"                            '��ư���� WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    btn.u1 = "http://www.linkhub.co.kr"     '�۸�ũ-iOS, ����ũ-Mobile
    btn.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    Buttons.Add btn

    Set btn = New PBKakaoButton
    
    btn.n = "�޽�������"
    btn.t = "MD"
    
    Buttons.Add btn
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, snd, "", "", altSendType, txtSndDT.Text, adsYN, Messages, Buttons, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================================
' �̹����� ÷�ε� 1���� ģ���� ������ �˺��� �����մϴ�.
' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - ��ü������ ���, ���乮��(MMS) ������ �����ϰ� ���� �ʽ��ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendFMS
'=========================================================================================
Private Sub btnSendFMS_ONE_Click()
    Dim receiptNum As String
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
    Dim requestNum As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '÷���̹��� ���ϰ��
    filePath = CommonDialog1.FileName
    
    '�̹��� ��ũ URL
    imageURL = "http://www.popbill.com"
    
    'īī���� ä�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042991"
    
    '��ü���� ��������, ����-������, C-ģ���峻�� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '���Ź�ȣ
    rcvInfo.rcv = "010000111"
    
    '�����ڸ�
    rcvInfo.rcvnm = "�������̸�"
    
    'ģ���� ����, �ִ� 400��
    rcvInfo.msg = "ģ���� �����Դϴ�. �̹��� ������ �����ϴ� ��� ģ���� ���ڼ��� �ִ� 400�� �Դϴ�."
    
    '��ü���� ����, ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
    rcvInfo.altsjt = "ģ���� �̹��� ��ü ���� ����"
    
    '��ü���� �޽��� ����
    rcvInfo.altmsg = "��ü���� �׽�Ʈ�Դϴ�."
    
    rcvList.Add rcvInfo
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"                        '��ư��
    btnInfo.t = "WL"                            '��ư���� DS-�����ȸ, WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    btnInfo.u1 = "http://www.linkhub.co.kr"     '�۸�ũ-iOS, ����ũ-Mobile
    btnInfo.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    
    receiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtSndDT.Text, adsYN, rcvList, btnList, filePath, imageURL, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum

End Sub

'=========================================================================================
' �̹����� ÷�ε� �ټ����� ģ���� ������ �˺��� �����ϸ�, ��� �����ڿ��� ���� ������ �����մϴ�. (�ִ� 1,000��)
' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - ��ü������ ���, ���乮��(MMS) ������ �����ϰ� ���� �ʽ��ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendFMS
'=========================================================================================
Private Sub btnSendFMS_SAME_Click()
    Dim receiptNum As String
    Dim plusFriendID As String
    Dim senderNum As String
    Dim altSendType As String
    Dim content As String
    Dim altSubject As String
    Dim altContent As String
    Dim adsYN As Boolean
    Dim filePath As String
    Dim i As Integer
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim btnList As New Collection
    Dim btnInfo As New PBKakaoButton
    Dim imageURL As String
    Dim requestNum As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '÷���̹��� ���ϰ��
    filePath = CommonDialog1.FileName
    
    '�̹��� ��ũ URL
    imageURL = "http://www.popbill.com"
    
    'īī���� ä�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042992"
    
    '�������� ����
    adsYN = True
    
    '(����) ģ���� ����, �ִ� 400��
    content = "ģ���� ���� �����Դϴ�. ������ ������ ���� �����ڿ��� �����ϴ� �����Դϴ�."
    
    '(����) ��ü���� ����
    ' ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
    altSubject = "ģ���� �̹��� ��ü���� ����"
    
    '(����) ��ü���� ����
    altContent = "��ü���� �׽�Ʈ�Դϴ�."
    
    '��ü���� ��������, ����-������, C-ģ���� ���� ����, A-��ü���ڳ��� ����
    altSendType = ""
   
    '�������� �迭, �ִ� 1000��
    For i = 0 To 10
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"                        '��ư��
    btnInfo.t = "WL"                            '��ư���� DS-�����ȸ, WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    btnInfo.u1 = "http://www.linkhub.co.kr"     '�۸�ũ-iOS, ����ũ-Mobile
    btnInfo.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, content, altContent, altSendType, txtSndDT.Text, adsYN, rcvList, btnList, filePath, imageURL, txtUserID.Text, requestNum, altSubject)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================================
' �̹����� ÷�ε� �ټ����� ģ���� ������ �˺��� �����ϸ�, ������ ���� ���� ������ �����մϴ�. (�ִ� 1,000��)
' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
' - ��ü������ ���, ���乮��(MMS) ������ �����ϰ� ���� �ʽ��ϴ�.
' - https://docs.popbill.com/kakao/vb/api#SendFMS
'=========================================================================================
Private Sub btnSendFMS_MULTI_Click()
    Dim receiptNum As String
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
    Dim requestNum As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '÷���̹��� ���ϰ��
    filePath = CommonDialog1.FileName
    
    '�̹��� ��ũ URL
    imageURL = "http://www.popbill.com"
    
    'īī���� ä�� ���̵�
    plusFriendID = "@�˺�"
    
    '�˺��� ���� ��ϵ� �߽Ź�ȣ
    senderNum = "07043042991"
    
    '��ü���� ��������, ����-������, C-ģ���� ���� ����, A-��ü���ڳ��� ����
    altSendType = ""
    
    '�������� ����
    adsYN = True
    
       
    '�������� �迭, �ִ� 1000��
    For i = 0 To 10
    
        Set rcvInfo = New PBKakaoReceiver
        
        '���Ź�ȣ
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '�����ڸ�
        rcvInfo.rcvnm = "������ �̸�" + CStr(i)
                 
        'ģ���� ����, �ִ� 400��
        rcvInfo.msg = "ģ���� �����Դϴ�. �����ڿ� ���� �ٸ� ������ �����մϴ�." + CStr(i)
        
        '��ü���� ����, ��ü���� ����(90byte)�� ���� �幮(LMS)�� ��� ����
        rcvInfo.altsjt = "ģ���� ��ü ���� ����"
        
        '��ü���� �޽��� ����
        rcvInfo.altmsg = "��ü���� �����Դϴ�. �����ڿ� ���� �ٸ� ������ ������ �� �ֽ��ϴ�." + CStr(i)
        
        rcvList.Add rcvInfo
    
    Next
    
    
    '��ư ��������, �ִ� 5������ �迭�� �߰� ����
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "��ư��"                        '��ư��
    btnInfo.t = "WL"                            '��ư���� DS-�����ȸ, WL-����ũ, AL-�۸�ũ, MD-�޽������� BK-��Ű����
    btnInfo.u1 = "http://www.linkhub.co.kr"     '�۸�ũ-iOS, ����ũ-Mobile
    btnInfo.u2 = "http://www.popbill.com"       '�۸�ũ-Android, ����ũ-PC
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "�޽�������"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtSndDT.Text, adsYN, rcvList, btnList, filePath, imageURL, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum


'=========================================================================
' �˺����� ��ȯ���� ������ȣ�� ���� �˸���/ģ���� ���ۻ��� �� ����� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetMessages
'=========================================================================
Private Sub btnGetMessages_Click()
    Dim tmp As String
    
    Dim sentInfo As PBKakaoSentResult
    Dim info As PBKakaoSentDetail
    Dim btnInfo As PBKakaoButton
    
    Set sentInfo = KakaoService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
     
    If sentInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "==== ���ۿ�û ������� ==== " + vbCrLf
    tmp = tmp + "contentType (īī���� ����) : " + sentInfo.contentType + vbCrLf
    tmp = tmp + "templateCode (���ø� �ڵ�) : " + sentInfo.templateCode + vbCrLf
    tmp = tmp + "plusFriendID (īī���� ä�� ���̵�) : " + sentInfo.plusFriendID + vbCrLf
    tmp = tmp + "sendNum (�߽Ź�ȣ) : " + sentInfo.sendNum + vbCrLf
    tmp = tmp + "altSubject (��ü���� ����) : " + sentInfo.altSubject + vbCrLf
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
    
    If (sentInfo.btns Is Nothing) = False Then
        tmp = tmp + "==== ��ư����====" + vbCrLf
        For Each btnInfo In sentInfo.btns
             tmp = tmp + "n (��ư��) : " + btnInfo.n + vbCrLf
             tmp = tmp + "t (��ư����) : " + btnInfo.t + vbCrLf
             tmp = tmp + "u1 (��ư��ũ1) : " + btnInfo.u1 + vbCrLf
             tmp = tmp + "u2 (��ư��ũ2) : " + btnInfo.u2 + vbCrLf + vbCrLf
        Next
    End If
     
    MsgBox (tmp)
    
    tmp = "====================== ���۰������ ======================" + vbCrLf
    tmp = tmp + "state(���ۻ��� �ڵ�) | sendDT(�����Ͻ�) | receiveNum(���Ź�ȣ) |  receiveName(�����ڸ�) | content(�˸���/ģ���� ����) | " + vbCrLf
    tmp = tmp + "result(���۰�� �ڵ�) | resultDT(���۰�� �����Ͻ�) | altSubject(��ü���� ����) | altContent(��ü���� ����) | altContentType(��ü���� ��������) | altSendDT(��ü���� �����Ͻ�) | "
    tmp = tmp + "altResult(��ü���� ���۰�� �ڵ�) | altResultDT(��ü���� ���۰�� �����Ͻ�) | receiptNum(������ȣ) | requestNum(��û��ȣ) | interOPrefKey (��Ʈ�� ����Ű)" + vbCrLf
    
    For Each info In sentInfo.msgs
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.content + " | "
        tmp = tmp + CStr(info.result) + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.altSubject + " | "
        tmp = tmp + info.altContent + " | "
        tmp = tmp + CStr(info.altContentType) + " | "
        tmp = tmp + info.altSendDT + " | "
        tmp = tmp + CStr(info.altResult) + " | "
        tmp = tmp + info.altResultDT + " | "
        tmp = tmp + info.receiptNum + " | "
        tmp = tmp + info.requestNum + " | "
        tmp = tmp + info.interOPRefKey
        tmp = tmp + vbCrLf
    Next
        
    txtResult.Text = tmp
End Sub

'=========================================================================
' �˺����� ��ȯ���� ������ȣ�� ���� ���������� īī������ ���� ����մϴ�. (����ð� 10�� ������ ����)
' - https://docs.popbill.com/kakao/vb/api#CancelReserve
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
' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� �˸���/ģ���� ���ۻ��� �� ����� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetMessagesRN
'=========================================================================
Private Sub btnGetMessagesRN_Click()
    Dim tmp As String
    Dim sentInfo As PBKakaoSentResult
    Dim info As PBKakaoSentDetail
    Dim btnInfo As PBKakaoButton
    
    Set sentInfo = KakaoService.GetMessagesRN(txtCorpNum.Text, txtRequestNum.Text)
     
    If sentInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "==== ���ۿ�û ������� ==== " + vbCrLf
    tmp = tmp + "contentType (īī���� ����) : " + sentInfo.contentType + vbCrLf
    tmp = tmp + "templateCode (���ø� �ڵ�) : " + sentInfo.templateCode + vbCrLf
    tmp = tmp + "plusFriendID (īī���� ä�� ���̵�) : " + sentInfo.plusFriendID + vbCrLf
    tmp = tmp + "sendNum (�߽Ź�ȣ) : " + sentInfo.sendNum + vbCrLf
    tmp = tmp + "altSubject (��ü���� ����) : " + sentInfo.altSubject + vbCrLf
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
    
    If (sentInfo.btns Is Nothing) = False Then
        tmp = tmp + "==== ��ư����====" + vbCrLf
        For Each btnInfo In sentInfo.btns
             tmp = tmp + "n (��ư��) : " + btnInfo.n + vbCrLf
             tmp = tmp + "t (��ư����) : " + btnInfo.t + vbCrLf
             tmp = tmp + "u1 (��ư��ũ1) : " + btnInfo.u1 + vbCrLf
             tmp = tmp + "u2 (��ư��ũ2) : " + btnInfo.u2 + vbCrLf + vbCrLf
        Next
    End If
    
    MsgBox (tmp)
    
    tmp = "====================== ���۰������ ======================" + vbCrLf
    tmp = tmp + "state(���ۻ��� �ڵ�) | sendDT(�����Ͻ�) | receiveNum(���Ź�ȣ) |  receiveName(�����ڸ�) | content(�˸���/ģ���� ����) | " + vbCrLf
    tmp = tmp + "result(���۰�� �ڵ�) | resultDT(���۰�� �����Ͻ�) | altSubject(��ü���� ����) | altContent(��ü���� ����) | altContentType(��ü���� ��������) | altSendDT(��ü���� �����Ͻ�) | "
    tmp = tmp + "altResult(��ü���� ���۰�� �ڵ�) | altResultDT(��ü���� ���۰�� �����Ͻ�) | receiptNum(������ȣ) | requestNum(��û��ȣ)" + vbCrLf
    
    For Each info In sentInfo.msgs
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.content + " | "
        tmp = tmp + CStr(info.result) + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.altSubject + " | "
        tmp = tmp + info.altContent + " | "
        tmp = tmp + CStr(info.altContentType) + " | "
        tmp = tmp + info.altSendDT + " | "
        tmp = tmp + CStr(info.altResult) + " | "
        tmp = tmp + info.altResultDT + " | "
        tmp = tmp + info.receiptNum + " | "
        tmp = tmp + info.requestNum
        tmp = tmp + vbCrLf
    Next
        
    txtResult.Text = tmp
End Sub

'=========================================================================
' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� ���������� īī������ ���� ����մϴ�. (����ð� 10�� ������ ����)
' - https://docs.popbill.com/kakao/vb/api#CancelReserveRN
'=========================================================================
Private Sub btnCancelReserveRN_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CancelReserveRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' īī���� ä���� ����ϰ� ������ Ȯ���ϴ� īī���� ä�� ���� ������ �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetPlusFriendMgtURL
'=========================================================================
Private Sub btnGetPlusFriendMgtURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetPlusFriendMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˺��� ����� ����ȸ���� īī���� ä�� ����� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#ListPlusFriendID
'=========================================================================
Private Sub btnListPlusFriendID_Click()
    Dim PlusFriendIDList As Collection
    Dim tmp As String
    Dim info As PBPlusFriend
    
    Set PlusFriendIDList = KakaoService.ListPlusFriendID(txtCorpNum.Text)
    
    If PlusFriendIDList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    For Each info In PlusFriendIDList
        tmp = tmp + "plusFriendID (īī���� ä�� ���̵�) : " + info.plusFriendID + vbCrLf
        tmp = tmp + "plusFriendName (īī���� ä�� �̸�) : " + info.plusFriendName + vbCrLf
        tmp = tmp + "regDT (����Ͻ�) : " + info.regDT + vbCrLf
        tmp = tmp + "state (ä�� ����) : " + CStr(info.state) + vbCrLf
        tmp = tmp + "stateDT (ä�� ���� �Ͻ�) : " + info.stateDT + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' �˸��� ���ø��� ��û�ϰ� ���νɻ� ����� Ȯ���ϸ� ��� ������ Ȯ���ϴ� �˸��� ���ø� ���� ������ �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
'=========================================================================
Private Sub btnGetATSTemplateMgtURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetATSTemplateMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ε� �˸��� ���ø� ������ Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetATSTemplate
'=========================================================================
Private Sub btnGetATSTemplate_Click()
    Dim template As PBATSTemplate
    Dim btnInfo As PBKakaoButton
    
    Dim templateCode As String
    templateCode = "022010000188"
    
    Set template = KakaoService.GetATSTemplate(txtCorpNum.Text, templateCode)
    
    If template Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = tmp + "==== �˸��� ���ø� ====" + vbCrLf
    tmp = tmp + "templateCode (���ø� �ڵ�) : " + template.templateCode + vbCrLf
    tmp = tmp + "templateName (���ø� ����) : " + template.templateName + vbCrLf
    tmp = tmp + "plusFriendID (īī���� ä�� �˻��� ���̵�) : " + template.plusFriendID + vbCrLf + vbCrLf
    tmp = tmp + "template (���ø� ����) : " + template.template + vbCrLf
    tmp = tmp + "appendix (�ΰ��޽���) : " + template.appendix + vbCrLf
    tmp = tmp + "ads (����޽���) : " + template.ads + vbCrLf + vbCrLf
    
    If (template.btns Is Nothing) = False Then
        For Each btnInfo In template.btns
                tmp = tmp + " n (��ư��) : " + btnInfo.n + vbCrLf
                tmp = tmp + " t (��ư����) : " + btnInfo.t + vbCrLf
                tmp = tmp + " u1 (��ư��ũ1) : " + btnInfo.u1 + vbCrLf
                tmp = tmp + " u2 (��ư��ũ2) : " + btnInfo.u2 + vbCrLf + vbCrLf
            Next
    End If
    MsgBox tmp
End Sub

'=========================================================================
' ���ε� �˸��� ���ø� ����� Ȯ���մϴ�.
' - ��ȯ�׸��� ���ø��ڵ�(templateCode)�� �˸��� ���۽� ���˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#ListATSTemplate
'=========================================================================
Private Sub btnListATSTemplate_Click()
    Dim tmp As String
    Dim templateList As Collection
    Dim info As PBATSTemplate
    Dim btnInfo As PBKakaoButton

    Set templateList = KakaoService.ListATSTemplate(txtCorpNum.Text)
    
    If templateList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If

    For Each info In templateList
        tmp = tmp + "==== �˸��� ���ø� ====" + vbCrLf
        tmp = tmp + "templateCode (���ø� �ڵ�) : " + info.templateCode + vbCrLf
        tmp = tmp + "templateName (���ø� ����) : " + info.templateName + vbCrLf
        tmp = tmp + "template (���ø� ����) : " + info.template + vbCrLf + vbCrLf
        tmp = tmp + "plusFriendID (īī���� ä�� �˻��� ���̵�) : " + info.plusFriendID + vbCrLf + vbCrLf
        tmp = tmp + "appendix (�ΰ��޽���) : " + info.appendix + vbCrLf
        tmp = tmp + "ads (����޽���) : " + info.ads + vbCrLf + vbCrLf
   
        If (info.btns Is Nothing) = False Then
            For Each btnInfo In info.btns
                tmp = tmp + " n (��ư��) : " + btnInfo.n + vbCrLf
                tmp = tmp + " t (��ư����) : " + btnInfo.t + vbCrLf
                tmp = tmp + " u1 (��ư��ũ1) : " + btnInfo.u1 + vbCrLf
                tmp = tmp + " u2 (��ư��ũ2) : " + btnInfo.u2 + vbCrLf + vbCrLf
            Next
        End If
     
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' īī���� �߽Ź�ȣ ��Ͽ��θ� Ȯ���մϴ�.
' - �߽Ź�ȣ ���°� '����'�� ��쿡�� ���ϰ� 'Response'�� ���� 'code'�� 1�� ��ȯ�˴ϴ�.
' - https://docs.popbill.com/kakao/vb/api#CheckSenderNumber
'=========================================================================
Private Sub btnCheckSenderNumber_Click()
    Dim Response As PBResponse
    Dim SenderNumber As String
    
    SenderNumber = "070-4304-2991"
    
    Set Response = KakaoService.CheckSenderNumber(txtCorpNum.Text, SenderNumber, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �߽Ź�ȣ�� ����ϰ� ������ Ȯ���ϴ� īī���� �߽Ź�ȣ ���� ������ �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetSenderNumberMgtURL
'=========================================================================
Private Sub btnGetSenderNumberMgtURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetSenderNumberMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If

    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˺��� ����� ����ȸ���� īī���� �߽Ź�ȣ ����� Ȯ���մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetSenderNumberList
'=========================================================================
Private Sub btnGetSenderNumberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim info As PBKakaoSenderNumber
    
    Set SenderNumberList = KakaoService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    For Each info In SenderNumberList
        tmp = tmp + "number (�߽Ź�ȣ) : " + info.number + vbCrLf
        tmp = tmp + "representYN (��ǥ��ȣ ��������) : " + CStr(info.representYN) + vbCrLf
        tmp = tmp + "state (��ϻ���) : " + CStr(info.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' īī���� ���۳����� Ȯ���ϴ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/kakao/vb/api#GetSentListURL
'=========================================================================
Private Sub btnGetSentListURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetSentListURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(KakaoService.LastErrCode) + vbCrLf + "����޽��� : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˻����ǿ� �ش��ϴ� īī���� ���۳����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 2����)
' - īī���� �����Ͻ÷κ��� 6���� �̳� �����Ǹ� ��ȸ�� �� �ֽ��ϴ�.
' - https://docs.popbill.com/kakao/vb/api#Search
'=========================================================================
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
    Dim QString As String
        
    '[�ʼ�] ��������, ��������(yyyyMMdd)
    SDate = "20220406"
    
    '[�ʼ�] ��������, ��������(yyyyMMdd)
    EDate = "20220406"
    
    '���ۻ��°� �迭 [0-���/ 1-������ / 2-���� / 3- ��ü / 4-���� / 5-���]
    state.Add "0"
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    state.Add "5"
    
    '�˻���� �迭  [ATS-�˸��� / FTS-ģ���� �ؽ�Ʈ / FMS-ģ���� �̹���]
    Item.Add "ATS"
    Item.Add "FTS"
    Item.Add "FMS"
    
    '���� �˸���/ģ���� �˻����� [����-��ü��ȸ / 1-�������� ��ȸ / 0-������� ��ȸ]
    ReserveYN = ""
    
    '������ȸ����, True(������ȸ), False(��ü��ȸ)
    SenderYN = False
    
    '������ ��ȣ, �⺻�� '1'
    Page = 1
    
    '������ ��ϰ���, �ִ� 1000��
    PerPage = 10
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    '��ȸ �˻��� ,�����ڸ� ����
    QString = ""
    

    Set searchList = KakaoService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order, QString)
     
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
    
    MsgBox (tmp)
    
    tmp = "====================== ���۰������ ======================" + vbCrLf
    tmp = tmp + "state(���ۻ��� �ڵ�) | sendDT(�����Ͻ�) | receiveNum(���Ź�ȣ) |  receiveName(�����ڸ�) | content(�˸���/ģ���� ����) | " + vbCrLf
    tmp = tmp + "result(���۰�� �ڵ�) | resultDT(���۰�� �����Ͻ�) | altContent(��ü���� ����) | altContent(��ü���� ����) | altContentType(��ü���� ��������) | altSendDT(��ü���� �����Ͻ�) | "
    tmp = tmp + "altResult(��ü���� ���۰�� �ڵ�) | altResultDT(��ü���� ���۰�� �����Ͻ�) | receiptNum(������ȣ) | requestNum(��û��ȣ)"
    
    For Each info In searchList.list
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.content + " | "
        tmp = tmp + CStr(info.result) + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.altSubject + " | "
        tmp = tmp + info.altContent + " | "
        tmp = tmp + CStr(info.altContentType) + " | "
        tmp = tmp + info.altSendDT + " | "
        tmp = tmp + CStr(info.altResult) + " | "
        tmp = tmp + info.altResultDT + " | "
        tmp = tmp + info.receiptNum + " | "
        tmp = tmp + info.requestNum
        tmp = tmp + vbCrLf
    Next
    
    txtResult.Text = tmp
    
End Sub

Private Sub Form_Load()

    'īī���� ��� �ʱ�ȭ
    KakaoService.Initialize linkID, SecretKey
    
    '����ȯ�漳����, True-���߿� False-�����
    KakaoService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
    KakaoService.IPRestrictOnOff = True
    
    '���ýý��� �ð� ��뿩�� True-���, False-�̻��, �⺻��(False)
    KakaoService.UseLocalTimeYN = False
End Sub





