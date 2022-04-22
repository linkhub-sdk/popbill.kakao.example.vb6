VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 카카오톡 API SDK VB6 Example"
   ClientHeight    =   13350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   ScaleHeight     =   13350
   ScaleWidth      =   18645
   StartUpPosition =   2  '화면 가운데
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
      Caption         =   " 팝빌 기본 API "
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   17760
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         ClipControls    =   0   'False
         Height          =   2415
         Left            =   11880
         TabIndex        =   21
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "전송단가"
         Height          =   2415
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   4920
         Begin VB.CommandButton btnGetChargeInfo_FMS 
            Caption         =   "친구톡 이미지 과금정보"
            Height          =   410
            Left            =   2520
            TabIndex        =   29
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_FTS 
            Caption         =   "친구톡 텍스트 과금정보"
            Height          =   410
            Left            =   2520
            TabIndex        =   28
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetUnitCost_FMS 
            Caption         =   "친구톡 이미지 전송단가"
            Height          =   410
            Left            =   150
            TabIndex        =   27
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetUnitCost_ATS 
            Caption         =   "알림톡 전송단가"
            Height          =   410
            Left            =   150
            TabIndex        =   20
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetUnitCost_FTS 
            Caption         =   "친구톡 텍스트 전송단가"
            Height          =   410
            Left            =   150
            TabIndex        =   19
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_ATS 
            Caption         =   "알림톡 과금정보"
            Height          =   410
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "담당자 관련"
         Height          =   2415
         Left            =   13800
         TabIndex        =   9
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "담당자 정보 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "회사정보 관련"
         Height          =   2415
         Left            =   15840
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "연동과금 포인트"
         Height          =   2415
         Left            =   6960
         TabIndex        =   4
         Top             =   240
         Width           =   2295
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "파트너과금 포인트"
         Height          =   2415
         Left            =   9360
         TabIndex        =   1
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "파트너 포인트 충전 URL"
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
      Caption         =   "팝빌 카카오톡 관련 기능"
      Height          =   9015
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   17775
      Begin VB.Frame Frame10 
         Caption         =   "요청번호 할당 전송건 처리"
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
            Caption         =   "전송상태 확인"
            Height          =   495
            Left            =   360
            TabIndex        =   63
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton btnCancelReserveRN 
            Caption         =   "예약전송 취소"
            Height          =   495
            Left            =   3120
            TabIndex        =   62
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "요청번호(requestNum) :"
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
         ScrollBars      =   3  '양방향
         TabIndex        =   57
         Top             =   5400
         Width           =   17295
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약전송 취소"
         Height          =   495
         Left            =   3120
         TabIndex        =   56
         Top             =   4560
         Width           =   2775
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "전송상태 확인"
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
         Caption         =   "카카오톡 관리"
         Height          =   3615
         Left            =   11760
         TabIndex        =   45
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton btnCheckSenderNumber 
            Caption         =   "발신번호 등록 여부 확인"
            Height          =   495
            Left            =   2880
            TabIndex        =   72
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton btnGetATSTemplate 
            Caption         =   "알림톡 템플릿 정보 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnSearch 
            Caption         =   "전송내역 목록 확인"
            Height          =   495
            Left            =   2880
            TabIndex        =   53
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSenderNumberMgtURL 
            Caption         =   "발신번호 관리 팝업 URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   52
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "발신번호 목록 확인"
            Height          =   495
            Left            =   2880
            TabIndex        =   51
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSentListURL 
            Caption         =   "전송내역 조회 팝업 URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   50
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnListATSTemplate 
            Caption         =   "알림톡 템플릿 목록 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CommandButton btnGetATSTemplateMgtURL 
            Caption         =   "알림톡 템플릿 관리 팝업 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnListPlusFriendID 
            Caption         =   "카카오톡 채널 목록 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetPlusFriendMgtURL 
            Caption         =   "카카오톡 채널 관리 팝업 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "친구톡 이미지 전송"
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   5415
         Begin VB.CommandButton btnSendFMS_multi 
            Caption         =   "대량 1000건 전송"
            Height          =   495
            Left            =   3480
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS_same 
            Caption         =   "동보 1000건 전송"
            Height          =   495
            Left            =   1680
            TabIndex        =   43
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS_ONE 
            Caption         =   "1건 전송"
            Height          =   495
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "친구톡 텍스트 전송"
         Height          =   855
         Left            =   5640
         TabIndex        =   37
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendFTS_one 
            Caption         =   "1건 전송"
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton btnSendFTS_same 
            Caption         =   "동보 1000건 전송"
            Height          =   495
            Left            =   1680
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFTS_multi 
            Caption         =   "대량 1000건 전송"
            Height          =   495
            Left            =   3480
            TabIndex        =   38
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "알림톡 전송"
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendATS_multi 
            Caption         =   "대량 1000건 전송"
            Height          =   495
            Left            =   3480
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_same 
            Caption         =   "동보 1000건 전송"
            Height          =   495
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_one 
            Caption         =   "1건 전송"
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
         Caption         =   "접수번호 관련 기능 (요청번호 미할당)"
         Height          =   1335
         Left            =   120
         TabIndex        =   58
         Top             =   3960
         Width           =   6255
         Begin VB.Label Label5 
            Caption         =   "접수번호(receiptNum) :"
            Height          =   375
            Left            =   180
            TabIndex        =   59
            Top             =   320
            Width           =   2175
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약시간(yyyyMMddHHmmss) : "
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
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   180
      Left            =   240
      TabIndex        =   26
      Top             =   360
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 아이디 : "
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
' 팝빌 카카오톡 API VB SDK Example
'
' - 업데이트 일자 : 2022-04-06
' - 연동 기술지원 연락처 : 1600-9854
' - 연동 기술지원 이메일 : code@linkhubcorp.com
' - VB SDK 적용방법 안내 : https://docs.popbill.com/kakao/tutorial/vb
'
' <테스트 연동개발 준비사항>
' 1) 30, 33번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 파트너 신청시 메일로 발급받은 인증정보를 참조하여 변경합니다.
'
' 2) 알림톡/친구톡을 전송하기 위해 발신번호 사전등록을 합니다. (등록방법은 사이트/API 두가지 방식이 있습니다.)
'   - 팝빌 사이트 로그인 > [문자/팩스] > [카카오톡] > [발신번호 사전등록] 메뉴에서 등록
'   - getSenderNumberMgtURL API를 통해 반환된 URL을 이용하여 발신번호 등록
'
' 3) 알림톡/친구톡을 전송하기 위해 카카오톡 채널를 등록 합니다. (등록방법은 사이트/API 두가지 방식이 있습니다.)
'   - 팝빌 사이트 로그인 > [문자/팩스] > [카카오톡] > [카카오톡 관리]  > 카카오톡 채널 계정관리 메뉴에서 등록
'   - GetPlusFriendMgtURL API를 통해 반환된 URL을 이용하여 카카오톡 채널 계정관리 등록
'
' 4) 알림톡 전송을 하기 위해  알림톡 템플릿을 신청 합니다.  (등록방법은 사이트/API 두가지 방식이 있습니다.)
'   - 팝빌 사이트 로그인 > [문자/팩스] > [카카오톡] > [카카오톡 관리]  > 알림톡 템플릿 관리 메뉴에서 등록
'   - GetATSTemplateMgtURL API를 통해 반환된 URL을 이용하여 알림톡 템플릿 등록
'=========================================================================

Option Explicit

'링크아이디
Private Const linkID = "TESTER"

'비밀키
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'카카오톡 서비스 클래스 선언
Private KakaoService As New PBKakaoService

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/kakao/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 이하
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 200자
    joinData.corpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 100자
    joinData.bizType = "업태"
    
    '종목, 최대 100자
    joinData.bizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    Set Response = KakaoService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 카카오톡(알림톡) 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_ATS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, ATS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "알림톡(ATS) 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 카카오톡(친구톡 텍스트) 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_FTS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FTS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "친구톡 텍스트(FTS) 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 카카오톡(친구톡 이미지) 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_FMS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "친구톡 이미지(FMS) 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 팝빌 카카오톡(알림톡) API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_ATS_Click()
    Dim ChargeInfo As PBChargeInfo

    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, ATS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "전송단가 (unitCost) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "과금유형 (chargeMethod) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "과금제도 (rateSystem) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 카카오톡(친구톡 텍스트) API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_FTS_Click()
    Dim ChargeInfo As PBChargeInfo
        
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FTS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "전송단가 (unitCost) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "과금유형 (chargeMethod) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "과금제도 (rateSystem) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 카카오톡(친구톡 이미지) API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_FMS_Click()
    Dim ChargeInfo As PBChargeInfo
        
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FMS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "전송단가 (unitCost) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "과금유형 (chargeMethod) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "과금제도 (rateSystem) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 (balance) : " + CStr(balance)
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()

    Dim URL As String
    
    URL = KakaoService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = KakaoService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = KakaoService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 (balance) : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
    
    URL = KakaoService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
    
    'Internet Explorer Browser 호출
    Dim IE As Object
   
    Dim strResult As String
    Dim strSiteName As String
   
    Set IE = CreateObject("InternetExplorer.Application")
    strSiteName = URL
    IE.Navigate strSiteName
    With IE
        .Visible = True     '브라우저창 활성화
        .Resizable = True   '브라우저창 크기 변경 On/Off
        .MenuBar = True     '메뉴바 On/Off
        .Toolbar = True     '툴바 On/Off
        .AddressBar = True  '주소바 On/Off
        .StatusBar = False  '상태바 On/Off
    End With
   
    Set IE = Nothing
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()

    Dim URL As String
    
    URL = KakaoService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
        
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/kakao/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 이하
    joinData.id = "testkorea"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf#$%123"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
        
    Set Response = KakaoService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub


'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = "testkorea"
    
    Set info = KakaoService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.tel + " | " _
            + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = KakaoService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If

    tmp = "id(아이디) | personName(성명) | email(이메일) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/kakao/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
                
    Set Response = KakaoService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = KakaoService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
            
    tmp = tmp + "대표자성명 (ceoname) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "상호명 (corpName) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "주소 (addr) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "업태 (bizType) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "종목 (bizClass) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다.
' - https://docs.popbill.com/kakao/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 200자
    CorpInfo.corpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.bizClass = "종목"
    
    Set Response = KakaoService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
'승인된 템플릿의 내용을 작성하여 1건의 알림톡 전송을 팝빌에 접수합니다.
' - 사전에 승인된 템플릿의 내용과 알림톡 전송내용(content)이 다를 경우 전송실패 처리됩니다.
' - 전송실패 시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
' - https://docs.popbill.com/kakao/vb/api#SendATS
'=========================================================================
Private Sub btnSendATS_ONE_Click()
    Dim templateCode As String
    Dim snd As String
    Dim content As String
    Dim altSendType As String
    Dim receiptNum As String
    Dim requestNum As String
    
    '알림톡 템플릿코드 - ListATStemplate API, GetPlusFriendMgtURL API, 또는 팝빌사이트에서 확인
    templateCode = "022040000005"
    
    '팝빌에 사전 등록된 발신번호
    snd = "07043042991"
    
    '알림톡 내용, 최대 1000자
    content = "[ 팝빌 ]" + vbCrLf
    content = content + "신청하신 #{템플릿코드}에 대한 심사가 완료되어 승인 처리되었습니다." + vbCrLf
    content = content + "해당 템플릿으로 전송 가능합니다." + vbCrLf + vbCrLf
    content = content + "문의사항 있으시면 파트너센터로 편하게 연락주시기 바랍니다." + vbCrLf + vbCrLf
    content = content + "팝빌 파트너센터 : 1600-8536" + vbCrLf
    content = content + "support@linkhub.co.kr"
    
    
    '대체문자 전송유형 (공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송)
    altSendType = "A"
    
    '카카오톡 전송정보
    Dim Messages As New Collection
    Dim info As New PBKakaoReceiver
    
    info.msg = content '알림톡 내용, 최대 1000자
    info.altsjt = "알림톡 대체 문자 제목"  '대체문자 제목, 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
    info.altmsg = "단건 알림톡 대체 문자 단건"  '대체문자 내용, 최대 2000byte
    info.rcv = "010111222"            '수신번호
    info.rcvnm = "popbill"            '수신자명
    info.interOPRefKey = ""  '파트너 지정키 (수신자 구분용)
    
    '수신자마다 다른내용의 버튼추가시 아래코드 참고
    'Set info.buttonList = New Collection
    'Dim detailButton As New PBKakaoButton
    'detailButton.n = "button"
    'detailButton.t = "WL"
    'detailButton.u1 = "test.popbill.com"
    'detailButton.u2 = "www.popbill.com"
    'info.buttonList.Add detailButton
     
    Messages.Add info
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    
    '알림톡 버튼정보를 템플릿 신청시 기재한 버튼정보와 동일하게 전송하는 경우 Buttons를 빈 배열로 처리.
    Dim Buttons As New Collection
    
    '알림톡 버튼 URL에 #{템플릿변수}를 기재한경우 템플릿변수 값을 변경하여 버튼정보 구성
    'Dim btn As PBKakaoButton
    
    'Set btn = New PBKakaoButton
    
    'btn.n = "버튼명"                        '버튼명
    'btn.t = "WL"                            '버튼유형 WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    'btn.u1 = "https://www.linkhub.co.kr"    '앱링크-iOS, 웹링크-Mobile
    'btn.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    'Buttons.Add btn
    
    
    receiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, snd, "", "", altSendType, txtSndDT.Text, Messages, txtUserID.Text, requestNum, Buttons, "")
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 승인된 템플릿 내용을 작성하여 다수건의 알림톡 전송을 팝빌에 접수하며, 모든 수신자에게 동일 내용을 전송합니다. (최대 1,000건)
' - 사전에 승인된 템플릿의 내용과 알림톡 전송내용(content)이 다를 경우 전송실패 처리됩니다.
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
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
    
    '알림톡 템플릿코드 - ListATStemplate API, GetPlusFriendMgtURL API, 또는 팝빌사이트에서 확인
    templateCode = "022040000005"
    
    '팝빌에 사전 등록된 발신번호
    snd = "07043042991"
    
    '(동보) 알림톡 내용, 최대 1000자
    content = "[ 팝빌 ]" + vbCrLf
    content = content + "신청하신 #{템플릿코드}에 대한 심사가 완료되어 승인 처리되었습니다." + vbCrLf
    content = content + "해당 템플릿으로 전송 가능합니다." + vbCrLf + vbCrLf
    content = content + "문의사항 있으시면 파트너센터로 편하게 연락주시기 바랍니다." + vbCrLf + vbCrLf
    content = content + "팝빌 파트너센터 : 1600-8536" + vbCrLf
    content = content + "support@linkhub.co.kr"
    
    '(동보) 대체문자 제목
    ' 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
    altSubject = "알림톡 대체문자 제목"
    
    '(동보) 대체문자 내용, 최대 2000byte
    altContent = "알림톡 대체 문자"
    
    '대체문자 전송유형 (공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송)
    altSendType = "A"
    
    '카카오톡 전송정보 배열, 최대 1000건
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver
    
    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)  '수신번호
        info.rcvnm = "popbill_" + CStr(i) '수신자명
        Messages.Add info
    Next
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '알림톡 버튼정보를 템플릿 신청시 기재한 버튼정보와 동일하게 전송하는 경우 Buttons를 빈 배열로 처리.
    Dim Buttons As New Collection
    
    '알림톡 버튼 URL에 #{템플릿변수}를 기재한경우 템플릿변수 값을 변경하여 버튼정보 구성
    'Dim btn As PBKakaoButton
    
    'Set btn = New PBKakaoButton
    
    'btn.n = "버튼명"                        '버튼명
    'btn.t = "WL"                            '버튼유형 WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    'btn.u1 = "https://www.linkhub.co.kr"    '앱링크-iOS, 웹링크-Mobile
    'btn.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    'Buttons.Add btn
    
    receiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, snd, content, altContent, altSendType, txtSndDT.Text, Messages, txtUserID.Text, requestNum, Buttons, altSubject)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum

End Sub

'=========================================================================
' 승인된 템플릿의 내용을 작성하여 다수건의 알림톡 전송을 팝빌에 접수하며, 수신자 별로 개별 내용을 전송합니다. (최대 1,000건)
' - 사전에 승인된 템플릿의 내용과 알림톡 전송내용(content)이 다를 경우 전송실패 처리됩니다.
' - 전송실패 시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
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
    
    '알림톡 템플릿코드 - ListATStemplate API, GetPlusFriendMgtURL API, 또는 팝빌사이트에서 확인
    templateCode = "022040000005"
    
    '팝빌에 사전 등록된 발신번호
    snd = "07043042991"
    
    '(동보) 알림톡 내용, 최대 1000자
    content = "[ 팝빌 ]" + vbCrLf
    content = content + "신청하신 #{템플릿코드}에 대한 심사가 완료되어 승인 처리되었습니다." + vbCrLf
    content = content + "해당 템플릿으로 전송 가능합니다." + vbCrLf + vbCrLf
    content = content + "문의사항 있으시면 파트너센터로 편하게 연락주시기 바랍니다." + vbCrLf + vbCrLf
    content = content + "팝빌 파트너센터 : 1600-8536" + vbCrLf
    content = content + "support@linkhub.co.kr"
    
    '대체문자 전송유형 (공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송)
    altSendType = "A"
    
    '카카오톡 수신정보 배열, 최대 1000건
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver

    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)                    '수신번호
        info.rcvnm = "popbill_" + CStr(i)                   '수신자명
        info.msg = content                                  '알림톡 내용, 최대 1000자
        info.altsjt = "알림톡 대체 문자 제목"               '대체문자 제목, 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
        info.altmsg = "알림톡 대량 대체 문자입니다." + CStr(i)   '대체문자 메시지 내용, 최대 2000byte
        Messages.Add info
    Next
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '알림톡 버튼정보를 템플릿 신청시 기재한 버튼정보와 동일하게 전송하는 경우 Buttons를 빈 배열로 처리.
    Dim Buttons As New Collection
    
    '알림톡 버튼 URL에 #{템플릿변수}를 기재한경우 템플릿변수 값을 변경하여 버튼정보 구성
    'Dim btn As PBKakaoButton
    
    'Set btn = New PBKakaoButton
    
    'btn.n = "버튼명"                        '버튼명
    'btn.t = "WL"                            '버튼유형 WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    'btn.u1 = "https://www.linkhub.co.kr"    '앱링크-iOS, 웹링크-Mobile
    'btn.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
   
    'Buttons.Add btn
    
    receiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, snd, "", "", altSendType, txtSndDT.Text, Messages, txtUserID.Text, requestNum, Buttons, "")
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 텍스트로 구성된 1건의 친구톡 전송을 팝빌에 접수합니다.
' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
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
    
    '카카오톡 채널 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    snd = "07043042991"
    
    '대체문자 전송유형 (공백-미전송 / C-친구톡내용 전송 / A-대체문자내용 전송)
    altSendType = "C"

    '광고전송 여부
    adsYN = False
    
    '카카오톡 메세지 구성
    Dim Messages As New Collection
    Dim info As New PBKakaoReceiver
    
    info.msg = "친구톡 텍스트 입니다"            '친구톡 내용, 최대 1000자
    info.altsjt = "친구톡 텍스트 대체 문자 제목" '대체문자 제목, 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
    info.altmsg = "친구톡 텍스트 대체 문자"      '대체문자 내용, 최대 2000byte
    info.rcv = "010000111"                       '수신번호
    info.rcvnm = "수신자이름"                    '수신자명
    
    Messages.Add info
        
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Dim Buttons As New Collection
    Dim btn As PBKakaoButton
    
    Set btn = New PBKakaoButton
    
    btn.n = "버튼명"                        '버튼명
    btn.t = "WL"                            '버튼유형 WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    btn.u1 = "http://www.linkhub.co.kr"     '앱링크-iOS, 웹링크-Mobile
    btn.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    Buttons.Add btn
    
    Set btn = New PBKakaoButton
    
    btn.n = "메시지전달"
    btn.t = "MD"
    
    Buttons.Add btn
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, snd, "", "", altSendType, txtSndDT.Text, adsYN, Messages, Buttons, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 텍스트로 구성된 다수건의 친구톡 전송을 팝빌에 접수하며, 모든 수신자에게 동일 내용을 전송합니다. (최대 1,000건)
' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
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
    
    '카카오톡 채널 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    snd = "07043042991"
    
    '(동보) 친구톡 내용, 최대 1000자
    content = "친구톡 텍스트 입니다"
    
    '(동보) 대체문자 제목
    ' 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
    altSubject = "친구톡 대체문자 제목"
    
    '(동보) 대체문자 내용, 최대 2000byte
    altContent = "친구톡 텍스트 대체 문자"
        
    '대체문자 전송유형 (공백-미전송 / C-친구톡내용 전송 / A-대체문자내용 전송)
    altSendType = "C"

    '광고전송 여부
    adsYN = False
    
    '수신정보 배열, 최대 1000건
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver
    
    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)    '수신번호
        info.rcvnm = "popbill_" + CStr(i)   '수신자명
        Messages.Add info
    Next
        
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Dim Buttons As New Collection
    Dim btn As PBKakaoButton
    
    Set btn = New PBKakaoButton
    
    btn.n = "버튼명"                        '버튼명
    btn.t = "WL"                            '버튼유형 WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    btn.u1 = "http://www.linkhub.co.kr"     '앱링크-iOS, 웹링크-Mobile
    btn.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    Buttons.Add btn
    
    Set btn = New PBKakaoButton
    
    btn.n = "메시지전달"
    btn.t = "MD"
    
    Buttons.Add btn
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, snd, content, altContent, altSendType, txtSndDT.Text, adsYN, Messages, Buttons, txtUserID.Text, requestNum, altSubject)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 텍스트로 구성된 다수건의 친구톡 전송을 팝빌에 접수하며, 수신자 별로 개별 내용을 전송합니다. (최대 1,000건)
' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
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
    
    '카카오톡 채널 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    snd = "07043042991"
    
    '대체문자 전송유형 (공백-미전송 / C-친구톡내용 전송 / A-대체문자내용 전송)
    altSendType = "A"

    '광고전송 여부
    adsYN = False
    
    '수신정보 배열, 최대 1000건
    Dim Messages As New Collection
    Dim info As PBKakaoReceiver
    
    For i = 1 To 10
        Set info = New PBKakaoReceiver
        info.rcv = "010123456" + CStr(i)                   '수신번호
        info.rcvnm = "popbill_" + CStr(i)                  '수신자명
        info.msg = "테스트 템플릿 입니다"                  '알림톡 내용, 최대 1000자
        info.altsjt = "친구톡 대체 문자 제목"              '대체문자 제목, 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
        info.altmsg = "친구톡 대체 문자입니다." + CStr(i)  '대체문자 메시지 내용, 최대 2000byte
        Messages.Add info
    Next
            
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Dim Buttons As New Collection
    Dim btn As PBKakaoButton
    
    Set btn = New PBKakaoButton
    
    btn.n = "버튼명"                        '버튼명
    btn.t = "WL"                            '버튼유형 WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    btn.u1 = "http://www.linkhub.co.kr"     '앱링크-iOS, 웹링크-Mobile
    btn.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    Buttons.Add btn

    Set btn = New PBKakaoButton
    
    btn.n = "메시지전달"
    btn.t = "MD"
    
    Buttons.Add btn
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, snd, "", "", altSendType, txtSndDT.Text, adsYN, Messages, Buttons, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================================
' 이미지가 첨부된 1건의 친구톡 전송을 팝빌에 접수합니다.
' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
' - 대체문자의 경우, 포토문자(MMS) 형식은 지원하고 있지 않습니다.
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
    
    '첨부이미지 파일경로
    filePath = CommonDialog1.FileName
    
    '이미지 링크 URL
    imageURL = "http://www.popbill.com"
    
    '카카오톡 채널 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042991"
    
    '대체문자 전송유형, 공백-미전송, C-친구톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '수신번호
    rcvInfo.rcv = "010000111"
    
    '수신자명
    rcvInfo.rcvnm = "수신자이름"
    
    '친구톡 내용, 최대 400자
    rcvInfo.msg = "친구톡 내용입니다. 이미지 파일을 전송하는 경우 친구톡 글자수는 최대 400자 입니다."
    
    '대체문자 제목, 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
    rcvInfo.altsjt = "친구톡 이미지 대체 문자 제목"
    
    '대체문자 메시지 내용
    rcvInfo.altmsg = "대체문자 테스트입니다."
    
    rcvList.Add rcvInfo
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"                        '버튼명
    btnInfo.t = "WL"                            '버튼유형 DS-배송조회, WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    btnInfo.u1 = "http://www.linkhub.co.kr"     '앱링크-iOS, 웹링크-Mobile
    btnInfo.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    
    receiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtSndDT.Text, adsYN, rcvList, btnList, filePath, imageURL, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum

End Sub

'=========================================================================================
' 이미지가 첨부된 다수건의 친구톡 전송을 팝빌에 접수하며, 모든 수신자에게 동일 내용을 전송합니다. (최대 1,000건)
' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
' - 대체문자의 경우, 포토문자(MMS) 형식은 지원하고 있지 않습니다.
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
    
    '첨부이미지 파일경로
    filePath = CommonDialog1.FileName
    
    '이미지 링크 URL
    imageURL = "http://www.popbill.com"
    
    '카카오톡 채널 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042992"
    
    '광고전송 여부
    adsYN = True
    
    '(동보) 친구톡 내용, 최대 400자
    content = "친구톡 전송 내용입니다. 동일한 내용을 개별 수신자에게 전송하는 예제입니다."
    
    '(동보) 대체문자 제목
    ' 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
    altSubject = "친구톡 이미지 대체문자 제목"
    
    '(동보) 대체문자 내용
    altContent = "대체문자 테스트입니다."
    
    '대체문자 전송유형, 공백-미전송, C-친구톡 내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 10
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"                        '버튼명
    btnInfo.t = "WL"                            '버튼유형 DS-배송조회, WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    btnInfo.u1 = "http://www.linkhub.co.kr"     '앱링크-iOS, 웹링크-Mobile
    btnInfo.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, content, altContent, altSendType, txtSndDT.Text, adsYN, rcvList, btnList, filePath, imageURL, txtUserID.Text, requestNum, altSubject)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================================
' 이미지가 첨부된 다수건의 친구톡 전송을 팝빌에 접수하며, 수신자 별로 개별 내용을 전송합니다. (최대 1,000건)
' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
' - 대체문자의 경우, 포토문자(MMS) 형식은 지원하고 있지 않습니다.
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
    
    '첨부이미지 파일경로
    filePath = CommonDialog1.FileName
    
    '이미지 링크 URL
    imageURL = "http://www.popbill.com"
    
    '카카오톡 채널 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042991"
    
    '대체문자 전송유형, 공백-미전송, C-친구톡 내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
       
    '수신정보 배열, 최대 1000건
    For i = 0 To 10
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
                 
        '친구톡 내용, 최대 400자
        rcvInfo.msg = "친구톡 내용입니다. 수신자에 따라 다른 내용을 전송합니다." + CStr(i)
        
        '대체문자 제목, 대체문자 길이(90byte)에 따라 장문(LMS)인 경우 적용
        rcvInfo.altsjt = "친구톡 대체 문자 제목"
        
        '대체문자 메시지 내용
        rcvInfo.altmsg = "대체문자 내용입니다. 수신자에 따라 다른 내용을 전송할 수 있습니다." + CStr(i)
        
        rcvList.Add rcvInfo
    
    Next
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"                        '버튼명
    btnInfo.t = "WL"                            '버튼유형 DS-배송조회, WL-웹링크, AL-앱링크, MD-메시지전달 BK-봇키워드
    btnInfo.u1 = "http://www.linkhub.co.kr"     '앱링크-iOS, 웹링크-Mobile
    btnInfo.u2 = "http://www.popbill.com"       '앱링크-Android, 웹링크-PC
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtSndDT.Text, adsYN, rcvList, btnList, filePath, imageURL, txtUserID.Text, requestNum, "")
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum


'=========================================================================
' 팝빌에서 반환받은 접수번호를 통해 알림톡/친구톡 전송상태 및 결과를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetMessages
'=========================================================================
Private Sub btnGetMessages_Click()
    Dim tmp As String
    
    Dim sentInfo As PBKakaoSentResult
    Dim info As PBKakaoSentDetail
    Dim btnInfo As PBKakaoButton
    
    Set sentInfo = KakaoService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
     
    If sentInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "==== 전송요청 요약정보 ==== " + vbCrLf
    tmp = tmp + "contentType (카카오톡 유형) : " + sentInfo.contentType + vbCrLf
    tmp = tmp + "templateCode (템플릿 코드) : " + sentInfo.templateCode + vbCrLf
    tmp = tmp + "plusFriendID (카카오톡 채널 아이디) : " + sentInfo.plusFriendID + vbCrLf
    tmp = tmp + "sendNum (발신번호) : " + sentInfo.sendNum + vbCrLf
    tmp = tmp + "altSubject (대체문자 제목) : " + sentInfo.altSubject + vbCrLf
    tmp = tmp + "altContent (대체문자 내용) : " + sentInfo.altContent + vbCrLf
    tmp = tmp + "altSendType (대체문자 유형) : " + sentInfo.altSendType + vbCrLf
    tmp = tmp + "reserveDT (예약일시) : " + sentInfo.reserveDT + vbCrLf
    tmp = tmp + "adsYN (광고전송 여부) : " + CStr(sentInfo.adsYN) + vbCrLf
    tmp = tmp + "imageURL (친구톡 이미지 URL) : " + sentInfo.imageURL + vbCrLf
    tmp = tmp + "sendCnt (전송건수) : " + sentInfo.sendCnt + vbCrLf
    tmp = tmp + "successCnt (성공건수) : " + sentInfo.successCnt + vbCrLf
    tmp = tmp + "failCnt (실패건수) : " + sentInfo.failCnt + vbCrLf
    tmp = tmp + "altCnt (대체문자 건수) : " + sentInfo.altCnt + vbCrLf
    tmp = tmp + "cancelCnt (취소건수) : " + sentInfo.cancelCnt + vbCrLf + vbCrLf
    
    If (sentInfo.btns Is Nothing) = False Then
        tmp = tmp + "==== 버튼정보====" + vbCrLf
        For Each btnInfo In sentInfo.btns
             tmp = tmp + "n (버튼명) : " + btnInfo.n + vbCrLf
             tmp = tmp + "t (버튼유형) : " + btnInfo.t + vbCrLf
             tmp = tmp + "u1 (버튼링크1) : " + btnInfo.u1 + vbCrLf
             tmp = tmp + "u2 (버튼링크2) : " + btnInfo.u2 + vbCrLf + vbCrLf
        Next
    End If
     
    MsgBox (tmp)
    
    tmp = "====================== 전송결과정보 ======================" + vbCrLf
    tmp = tmp + "state(전송상태 코드) | sendDT(전송일시) | receiveNum(수신번호) |  receiveName(수신자명) | content(알림톡/친구톡 내용) | " + vbCrLf
    tmp = tmp + "result(전송결과 코드) | resultDT(전송결과 수신일시) | altSubject(대체문자 제목) | altContent(대체문자 내용) | altContentType(대체문자 전송유형) | altSendDT(대체문자 전송일시) | "
    tmp = tmp + "altResult(대체문자 전송결과 코드) | altResultDT(대체문자 전송결과 수신일시) | receiptNum(접수번호) | requestNum(요청번호) | interOPrefKey (파트너 지정키)" + vbCrLf
    
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
' 팝빌에서 반환받은 접수번호를 통해 예약접수된 카카오톡을 전송 취소합니다. (예약시간 10분 전까지 가능)
' - https://docs.popbill.com/kakao/vb/api#CancelReserve
'=========================================================================
Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 파트너가 할당한 전송요청 번호를 통해 알림톡/친구톡 전송상태 및 결과를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetMessagesRN
'=========================================================================
Private Sub btnGetMessagesRN_Click()
    Dim tmp As String
    Dim sentInfo As PBKakaoSentResult
    Dim info As PBKakaoSentDetail
    Dim btnInfo As PBKakaoButton
    
    Set sentInfo = KakaoService.GetMessagesRN(txtCorpNum.Text, txtRequestNum.Text)
     
    If sentInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "==== 전송요청 요약정보 ==== " + vbCrLf
    tmp = tmp + "contentType (카카오톡 유형) : " + sentInfo.contentType + vbCrLf
    tmp = tmp + "templateCode (템플릿 코드) : " + sentInfo.templateCode + vbCrLf
    tmp = tmp + "plusFriendID (카카오톡 채널 아이디) : " + sentInfo.plusFriendID + vbCrLf
    tmp = tmp + "sendNum (발신번호) : " + sentInfo.sendNum + vbCrLf
    tmp = tmp + "altSubject (대체문자 제목) : " + sentInfo.altSubject + vbCrLf
    tmp = tmp + "altContent (대체문자 내용) : " + sentInfo.altContent + vbCrLf
    tmp = tmp + "altSendType (대체문자 유형) : " + sentInfo.altSendType + vbCrLf
    tmp = tmp + "reserveDT (예약일시) : " + sentInfo.reserveDT + vbCrLf
    tmp = tmp + "adsYN (광고전송 여부) : " + CStr(sentInfo.adsYN) + vbCrLf
    tmp = tmp + "imageURL (친구톡 이미지 URL) : " + sentInfo.imageURL + vbCrLf
    tmp = tmp + "sendCnt (전송건수) : " + sentInfo.sendCnt + vbCrLf
    tmp = tmp + "successCnt (성공건수) : " + sentInfo.successCnt + vbCrLf
    tmp = tmp + "failCnt (실패건수) : " + sentInfo.failCnt + vbCrLf
    tmp = tmp + "altCnt (대체문자 건수) : " + sentInfo.altCnt + vbCrLf
    tmp = tmp + "cancelCnt (취소건수) : " + sentInfo.cancelCnt + vbCrLf + vbCrLf
    
    If (sentInfo.btns Is Nothing) = False Then
        tmp = tmp + "==== 버튼정보====" + vbCrLf
        For Each btnInfo In sentInfo.btns
             tmp = tmp + "n (버튼명) : " + btnInfo.n + vbCrLf
             tmp = tmp + "t (버튼유형) : " + btnInfo.t + vbCrLf
             tmp = tmp + "u1 (버튼링크1) : " + btnInfo.u1 + vbCrLf
             tmp = tmp + "u2 (버튼링크2) : " + btnInfo.u2 + vbCrLf + vbCrLf
        Next
    End If
    
    MsgBox (tmp)
    
    tmp = "====================== 전송결과정보 ======================" + vbCrLf
    tmp = tmp + "state(전송상태 코드) | sendDT(전송일시) | receiveNum(수신번호) |  receiveName(수신자명) | content(알림톡/친구톡 내용) | " + vbCrLf
    tmp = tmp + "result(전송결과 코드) | resultDT(전송결과 수신일시) | altSubject(대체문자 제목) | altContent(대체문자 내용) | altContentType(대체문자 전송유형) | altSendDT(대체문자 전송일시) | "
    tmp = tmp + "altResult(대체문자 전송결과 코드) | altResultDT(대체문자 전송결과 수신일시) | receiptNum(접수번호) | requestNum(요청번호)" + vbCrLf
    
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
' 파트너가 할당한 전송요청 번호를 통해 예약접수된 카카오톡을 전송 취소합니다. (예약시간 10분 전까지 가능)
' - https://docs.popbill.com/kakao/vb/api#CancelReserveRN
'=========================================================================
Private Sub btnCancelReserveRN_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CancelReserveRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 카카오톡 채널을 등록하고 내역을 확인하는 카카오톡 채널 관리 페이지 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetPlusFriendMgtURL
'=========================================================================
Private Sub btnGetPlusFriendMgtURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetPlusFriendMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 팝빌에 등록한 연동회원의 카카오톡 채널 목록을 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#ListPlusFriendID
'=========================================================================
Private Sub btnListPlusFriendID_Click()
    Dim PlusFriendIDList As Collection
    Dim tmp As String
    Dim info As PBPlusFriend
    
    Set PlusFriendIDList = KakaoService.ListPlusFriendID(txtCorpNum.Text)
    
    If PlusFriendIDList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    For Each info In PlusFriendIDList
        tmp = tmp + "plusFriendID (카카오톡 채널 아이디) : " + info.plusFriendID + vbCrLf
        tmp = tmp + "plusFriendName (카카오톡 채널 이름) : " + info.plusFriendName + vbCrLf
        tmp = tmp + "regDT (등록일시) : " + info.regDT + vbCrLf
        tmp = tmp + "state (채널 상태) : " + CStr(info.state) + vbCrLf
        tmp = tmp + "stateDT (채널 상태 일시) : " + info.stateDT + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 알림톡 템플릿을 신청하고 승인심사 결과를 확인하며 등록 내역을 확인하는 알림톡 템플릿 관리 페이지 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
'=========================================================================
Private Sub btnGetATSTemplateMgtURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetATSTemplateMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 승인된 알림톡 템플릿 정보를 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetATSTemplate
'=========================================================================
Private Sub btnGetATSTemplate_Click()
    Dim template As PBATSTemplate
    Dim btnInfo As PBKakaoButton
    
    Dim templateCode As String
    templateCode = "022010000188"
    
    Set template = KakaoService.GetATSTemplate(txtCorpNum.Text, templateCode)
    
    If template Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = tmp + "==== 알림톡 템플릿 ====" + vbCrLf
    tmp = tmp + "templateCode (템플릿 코드) : " + template.templateCode + vbCrLf
    tmp = tmp + "templateName (템플릿 제목) : " + template.templateName + vbCrLf
    tmp = tmp + "plusFriendID (카카오톡 채널 검색용 아이디) : " + template.plusFriendID + vbCrLf + vbCrLf
    tmp = tmp + "template (템플릿 내용) : " + template.template + vbCrLf
    tmp = tmp + "appendix (부가메시지) : " + template.appendix + vbCrLf
    tmp = tmp + "ads (광고메시지) : " + template.ads + vbCrLf + vbCrLf
    
    If (template.btns Is Nothing) = False Then
        For Each btnInfo In template.btns
                tmp = tmp + " n (버튼명) : " + btnInfo.n + vbCrLf
                tmp = tmp + " t (버튼유형) : " + btnInfo.t + vbCrLf
                tmp = tmp + " u1 (버튼링크1) : " + btnInfo.u1 + vbCrLf
                tmp = tmp + " u2 (버튼링크2) : " + btnInfo.u2 + vbCrLf + vbCrLf
            Next
    End If
    MsgBox tmp
End Sub

'=========================================================================
' 승인된 알림톡 템플릿 목록을 확인합니다.
' - 반환항목중 템플릿코드(templateCode)는 알림톡 전송시 사용됩니다.
' - https://docs.popbill.com/kakao/vb/api#ListATSTemplate
'=========================================================================
Private Sub btnListATSTemplate_Click()
    Dim tmp As String
    Dim templateList As Collection
    Dim info As PBATSTemplate
    Dim btnInfo As PBKakaoButton

    Set templateList = KakaoService.ListATSTemplate(txtCorpNum.Text)
    
    If templateList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If

    For Each info In templateList
        tmp = tmp + "==== 알림톡 템플릿 ====" + vbCrLf
        tmp = tmp + "templateCode (템플릿 코드) : " + info.templateCode + vbCrLf
        tmp = tmp + "templateName (템플릿 제목) : " + info.templateName + vbCrLf
        tmp = tmp + "template (템플릿 내용) : " + info.template + vbCrLf + vbCrLf
        tmp = tmp + "plusFriendID (카카오톡 채널 검색용 아이디) : " + info.plusFriendID + vbCrLf + vbCrLf
        tmp = tmp + "appendix (부가메시지) : " + info.appendix + vbCrLf
        tmp = tmp + "ads (광고메시지) : " + info.ads + vbCrLf + vbCrLf
   
        If (info.btns Is Nothing) = False Then
            For Each btnInfo In info.btns
                tmp = tmp + " n (버튼명) : " + btnInfo.n + vbCrLf
                tmp = tmp + " t (버튼유형) : " + btnInfo.t + vbCrLf
                tmp = tmp + " u1 (버튼링크1) : " + btnInfo.u1 + vbCrLf
                tmp = tmp + " u2 (버튼링크2) : " + btnInfo.u2 + vbCrLf + vbCrLf
            Next
        End If
     
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 카카오톡 발신번호 등록여부를 확인합니다.
' - 발신번호 상태가 '승인'인 경우에만 리턴값 'Response'의 변수 'code'가 1로 반환됩니다.
' - https://docs.popbill.com/kakao/vb/api#CheckSenderNumber
'=========================================================================
Private Sub btnCheckSenderNumber_Click()
    Dim Response As PBResponse
    Dim SenderNumber As String
    
    SenderNumber = "070-4304-2991"
    
    Set Response = KakaoService.CheckSenderNumber(txtCorpNum.Text, SenderNumber, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 발신번호를 등록하고 내역을 확인하는 카카오톡 발신번호 관리 페이지 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetSenderNumberMgtURL
'=========================================================================
Private Sub btnGetSenderNumberMgtURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetSenderNumberMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If

    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 팝빌에 등록한 연동회원의 카카오톡 발신번호 목록을 확인합니다.
' - https://docs.popbill.com/kakao/vb/api#GetSenderNumberList
'=========================================================================
Private Sub btnGetSenderNumberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim info As PBKakaoSenderNumber
    
    Set SenderNumberList = KakaoService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    For Each info In SenderNumberList
        tmp = tmp + "number (발신번호) : " + info.number + vbCrLf
        tmp = tmp + "representYN (대표번호 지정여부) : " + CStr(info.representYN) + vbCrLf
        tmp = tmp + "state (등록상태) : " + CStr(info.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 카카오톡 전송내역을 확인하는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/kakao/vb/api#GetSentListURL
'=========================================================================
Private Sub btnGetSentListURL_Click()
    Dim URL As String
    
    URL = KakaoService.GetSentListURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 검색조건에 해당하는 카카오톡 전송내역을 조회합니다. (조회기간 단위 : 최대 2개월)
' - 카카오톡 접수일시로부터 6개월 이내 접수건만 조회할 수 있습니다.
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
        
    '[필수] 시작일자, 날자형식(yyyyMMdd)
    SDate = "20220406"
    
    '[필수] 종료일자, 날자형식(yyyyMMdd)
    EDate = "20220406"
    
    '전송상태값 배열 [0-대기/ 1-전송중 / 2-성공 / 3- 대체 / 4-실패 / 5-취소]
    state.Add "0"
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    state.Add "5"
    
    '검색대상 배열  [ATS-알림톡 / FTS-친구톡 텍스트 / FMS-친구톡 이미지]
    Item.Add "ATS"
    Item.Add "FTS"
    Item.Add "FMS"
    
    '예약 알림톡/친구톡 검색여부 [공백-전체조회 / 1-예약전송 조회 / 0-즉시전송 조회]
    ReserveYN = ""
    
    '개인조회여부, True(개인조회), False(전체조회)
    SenderYN = False
    
    '페이지 번호, 기본값 '1'
    Page = 1
    
    '페이지 목록개수, 최대 1000건
    PerPage = 10
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '조회 검색어 ,수신자명 기재
    QString = ""
    

    Set searchList = KakaoService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order, QString)
     
    If searchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (응답코드) : " + CStr(searchList.code) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + searchList.message + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(searchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(searchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(searchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(searchList.pageCount) + vbCrLf + vbCrLf
    
    MsgBox (tmp)
    
    tmp = "====================== 전송결과정보 ======================" + vbCrLf
    tmp = tmp + "state(전송상태 코드) | sendDT(전송일시) | receiveNum(수신번호) |  receiveName(수신자명) | content(알림톡/친구톡 내용) | " + vbCrLf
    tmp = tmp + "result(전송결과 코드) | resultDT(전송결과 수신일시) | altContent(대체문자 내용) | altContent(대체문자 내용) | altContentType(대체문자 전송유형) | altSendDT(대체문자 전송일시) | "
    tmp = tmp + "altResult(대체문자 전송결과 코드) | altResultDT(대체문자 전송결과 수신일시) | receiptNum(접수번호) | requestNum(요청번호)"
    
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

    '카카오톡 모듈 초기화
    KakaoService.Initialize linkID, SecretKey
    
    '연동환경설정값, True-개발용 False-상업용
    KakaoService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-사용, False-미사용, 기본값(True)
    KakaoService.IPRestrictOnOff = True
    
    '로컬시스템 시간 사용여부 True-사용, False-미사용, 기본값(False)
    KakaoService.UseLocalTimeYN = False
End Sub





