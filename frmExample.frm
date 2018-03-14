VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 카카오톡 API SDK VB6 Example"
   ClientHeight    =   11415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17580
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   17580
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame4 
      Caption         =   "팝빌 카카오톡 관련 기능"
      Height          =   8055
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   17295
      Begin VB.TextBox txtResult 
         Height          =   4080
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   59
         Top             =   3480
         Width           =   16815
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약전송 취소"
         Height          =   495
         Left            =   6480
         TabIndex        =   58
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "전송상태 확인"
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
         Caption         =   "카카오톡 관리"
         Height          =   2895
         Left            =   11280
         TabIndex        =   46
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton btnSearch 
            Caption         =   "전송내역 목록 확인"
            Height          =   495
            Left            =   2880
            TabIndex        =   54
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_SENDER 
            Caption         =   "발신번호 관리 팝업 URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   53
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "발신번호 목록 확인"
            Height          =   495
            Left            =   2880
            TabIndex        =   52
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_BOX 
            Caption         =   "전송내역 조회 팝업 URL"
            Height          =   495
            Left            =   2880
            TabIndex        =   51
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnListATSTemplate 
            Caption         =   "알림톡 템플릿 목록 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   50
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_TEMPLATE 
            Caption         =   "알림톡 템플릿 관리 팝업 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton btnListPlusFriendID 
            Caption         =   "플러스친구 목록 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton btnGetURL_PLUSFRIEND 
            Caption         =   "플러스친구 계정관리 팝업 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "친구톡 이미지 전송"
         Height          =   855
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   5415
         Begin VB.CommandButton btnSendFMS_multi 
            Caption         =   "개별 1000건 전송"
            Height          =   495
            Left            =   3480
            TabIndex        =   45
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS_same 
            Caption         =   "대량 1000건 전송"
            Height          =   495
            Left            =   1680
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFMS 
            Caption         =   "1건 전송"
            Height          =   495
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "친구톡 텍스트 전송"
         Height          =   855
         Left            =   5640
         TabIndex        =   38
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendFTS_one 
            Caption         =   "1건 전송"
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton btnSendFTS_same 
            Caption         =   "대량 1000건 전송"
            Height          =   495
            Left            =   1680
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendFTS_multi 
            Caption         =   "개별 1000건 전송"
            Height          =   495
            Left            =   3480
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "알림톡 전송"
         Height          =   855
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton btnSendATS_multi 
            Caption         =   "개별 1000건 전송"
            Height          =   495
            Left            =   3480
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_same 
            Caption         =   "대량 1000건 전송"
            Height          =   495
            Left            =   1680
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btnSendATS_one 
            Caption         =   "1건 전송"
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
         Caption         =   "접수번호 : "
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약시간(yyyyMMddHHmmss) : "
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
      Caption         =   " 팝빌 기본 API "
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   17400
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         ClipControls    =   0   'False
         Height          =   1935
         Left            =   11520
         TabIndex        =   22
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "전송단가"
         Height          =   1935
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   4920
         Begin VB.CommandButton btnGetChargeInfo_FMS 
            Caption         =   "친구톡 이미지 과금정보"
            Height          =   410
            Left            =   2520
            TabIndex        =   30
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_FTS 
            Caption         =   "친구톡 텍스트 과금정보"
            Height          =   410
            Left            =   2520
            TabIndex        =   29
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost_FMS 
            Caption         =   "친구톡 이미지 전송단가"
            Height          =   410
            Left            =   150
            TabIndex        =   28
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost_ATS 
            Caption         =   "알림톡 전송단가"
            Height          =   410
            Left            =   150
            TabIndex        =   21
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost_FTS 
            Caption         =   "친구톡 텍스트 전송단가"
            Height          =   410
            Index           =   0
            Left            =   150
            TabIndex        =   20
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo_ATS 
            Caption         =   "알림톡 과금정보"
            Height          =   410
            Left            =   2520
            TabIndex        =   19
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "담당자 관련"
         Height          =   1935
         Left            =   13440
         TabIndex        =   10
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "회사정보 관련"
         Height          =   1935
         Left            =   15480
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "연동과금 포인트"
         Height          =   1935
         Left            =   6960
         TabIndex        =   4
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "파트너과금 포인트"
         Height          =   1935
         Left            =   9000
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 아이디 : "
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
' 팝빌 카카오톡 API VB 6.0 SDK Example
'
' - VB6 SDK 연동환경 설정방법 안내 : http://blog.linkhub.co.kr/569/
' - 업데이트 일자 : 2018-03-14
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' - 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
'
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const LinkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private KakaoService As New PBKakaoService

'=========================================================================
' 예약문자전송을 취소합니다.
' - 예약취소는 예약전송시간 10분전까지만 가능합니다.
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
' 팝빌 회원아이디 중복여부를 확인합니다.
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
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = KakaoService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 알림톡(ATS) 서비스 과금정보를 확인합니다.
'=========================================================================
Private Sub btnGetChargeInfo_ATS_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
            
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, ATS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (전송단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 친구톡 이미지(FMS) 서비스 과금정보를 확인합니다.
'=========================================================================
Private Sub btnGetChargeInfo_FMS_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FMS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (전송단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 친구톡 텍스트(FTS) 서비스 과금정보를 확인합니다.
'=========================================================================
Private Sub btnGetChargeInfo_FTS_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = KakaoService.GetChargeInfo(txtCorpNum.Text, FTS)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (전송단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = KakaoService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(대표자성명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(상호명) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetMessages_Click()
    Dim sentInfo As PBKakaoSentResult
    Dim tmp As String
    Dim info As PBKakaoSentDetail
    Dim btnInfo As PBKakaoButton
    
    Set sentInfo = KakaoService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
     
    If sentInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "================== 전송요청 요약정보 ================== " + vbCrLf
    tmp = tmp + "contentType (카카오톡 유형) : " + sentInfo.contentType + vbCrLf
    tmp = tmp + "templateCode (템플릿 코드) : " + sentInfo.templateCode + vbCrLf
    tmp = tmp + "plusFriendID (플러스친구 아이디) : " + sentInfo.plusFriendID + vbCrLf
    tmp = tmp + "sendNum (발신번호) : " + sentInfo.sendNum + vbCrLf
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
    
    tmp = tmp + "================== 버튼정보 ==================" + vbCrLf
    
    For Each btnInfo In sentInfo.btns
        tmp = tmp + "n (버튼명) : " + btnInfo.n + vbCrLf
        tmp = tmp + "t (버튼유형) : " + btnInfo.t + vbCrLf
        tmp = tmp + "u1 (버튼링크1) : " + btnInfo.u1 + vbCrLf
        tmp = tmp + "u2 (버튼링크2) : " + btnInfo.u2 + vbCrLf + vbCrLf
    Next
    
    
    tmp = tmp + vbCrLf + "================== 전송결과정보 ==================" + vbCrLf
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
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
'   이용하시기 바랍니다.
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = KakaoService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = KakaoService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원 포인트 충전 팝업 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = KakaoService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = KakaoService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
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
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    For Each SenderNumber In SenderNumberList
        tmp = tmp + "발신번호(number) : " + SenderNumber.number + vbCrLf
        tmp = tmp + "대표번호 지정여부(representYN) : " + CStr(SenderNumber.representYN) + vbCrLf
        tmp = tmp + "등록상태(state) : " + CStr(SenderNumber.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetURL_BOX_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_PLUSFRIEND_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "PLUSFRIEND")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_SENDER_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "SENDER")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_TEMPLATE_Click()
    Dim url As String
    
    url = KakaoService.GetURL(txtCorpNum.Text, txtUserID.Text, "TEMPLATE")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 연동회원 가입을 요청합니다.
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
 
    '링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호
    joinData.CorpNum = "0000000403"
    
    '대표자성명
    joinData.ceoname = "대표자성명"
    
    '상호명
    joinData.corpName = "회원상호"
    
    '주소
    joinData.addr = "주소"
    
    '업태
    joinData.bizType = "업태"
    
    '종목
    joinData.bizClass = "종목"
    
    '아이디
    joinData.id = "userid"
    
    '비밀번호
    joinData.pwd = "pwd_must_be_long_enough"
    
    '담당자명
    joinData.ContactName = "담당자성명"
    
    '담당자 연락처
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 메일
    joinData.ContactEmail = "test@test.com"
    
    Set Response = KakaoService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

Private Sub btnListATSTemplate_Click()
    
    Dim tmp As String
    Dim templateList As Collection
    Dim atsTemplate As PBATSTemplate
    Dim btnInfo As PBKakaoButton
    
    
    Set templateList = KakaoService.ListATSTemplate(txtCorpNum.Text)
    
    If templateList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    For Each atsTemplate In templateList
        tmp = tmp + "===================== 템플릿 정보 =====================" + vbCrLf
        tmp = tmp + "템플릿 코드(templateCode) : " + atsTemplate.templateCode + vbCrLf
        tmp = tmp + "템플릿 제목(templateName) : " + atsTemplate.templateName + vbCrLf
        tmp = tmp + "템플릿 내용(template) : " + atsTemplate.template + vbCrLf
        tmp = tmp + "플러스친구 아이디(plusFriendID) : " + atsTemplate.plusFriendID + vbCrLf + vbCrLf
        
        If (atsTemplate.btns Is Nothing) = False Then
            For Each btnInfo In atsTemplate.btns
                tmp = tmp + "버튼명(n) : " + btnInfo.n + vbCrLf
                tmp = tmp + "버튼유형(t) : " + btnInfo.t + vbCrLf
                tmp = tmp + "버튼링크1(u1) : " + btnInfo.u1 + vbCrLf
                tmp = tmp + "버튼링크2(u2) : " + btnInfo.u2 + vbCrLf + vbCrLf
            Next
        End If
        
        tmp = tmp + vbCrLf
        
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
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
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
        
    For Each plusFriendInfo In plusFriendList
        tmp = tmp + "플러스친구 아이디(plusFriendID) : " + plusFriendInfo.plusFriendID + vbCrLf
        tmp = tmp + "플러스친구 이름(plusFriendName) : " + plusFriendInfo.plusFriendName + vbCrLf
        tmp = tmp + "등록일시(regDT) : " + plusFriendInfo.regDT + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = "testkorea_20161011"
    
    '비밀번호
    joinData.pwd = "test@test.com"
    
    '담당자명
    joinData.personName = "담당자명"
    
    '담당자 연락처
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '담당자 메일주소
    joinData.email = "test@test.com"
        
    '회사조회 권한여부, true-회사조회 / false-개인조회
    joinData.searchAllAllowYN = True
            
    Set Response = KakaoService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
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
    
    '[필수] 시작일자, 날자형식(yyyyMMdd)
    SDate = "20180101"
    
    '[필수] 종료일자, 날자형식(yyyyMMdd)
    EDate = "20181231"
    
    '전송상태값 배열, 0-대기, 1-전송중, 2-성공, 3- 대체, 4-실패, 5-취소
    state.Add "0"
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    state.Add "5"
    
    '검색대상 배열, ATS(알림톡), FTS(친구톡 텍스트), FMS(친구톡 이미지)
    Item.Add "ATS"
    Item.Add "FTS"
    Item.Add "FMS"
    
    '예약문자 검색여부, 공백(전체조회), 1(예약전송 조회), 0(즉시전송 조회)
    ReserveYN = ""
    
    '개인조회여부, True(개인조회), False(전체조회)
    SenderYN = False
    
    '페이지 번호
    Page = 1
    
    '페이지 목록개수, 최대 1000건
    PerPage = 50
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"

    Set searchList = KakaoService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order)
     
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
    
    ' 알림톡 템플릿코드 - ListATStemplate API, GetURL(TEMPLATE) API, 또는 팝빌사이트에서 확인
    templateCode = "018020000001"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 99
        
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)

        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
        
        '알림톡 내용, 최대 1000자
        rcvInfo.msg = "[테스트] 테스트 템플릿입니다." + CStr(i)
        
        '대체문자 메시지 내용
        rcvInfo.altmsg = "대체문자 내용입니다." + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    ReceiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, senderNum, "", "", altSendType, txtReserveDT.Text, rcvList)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendATS_one_Click()
    Dim rcvList As New Collection
    Dim rcvInfo As New PBKakaoReceiver
    Dim ReceiptNum As String
    Dim templateCode As String
    Dim senderNum As String
    Dim altSendType As String
    
    '알림톡 템플릿코드 - ListATStemplate API, GetURL(TEMPLATE) API, 또는 팝빌사이트에서 확인
    templateCode = "018020000001"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '수신번호
    rcvInfo.rcv = "010000111"
    
    '수신자명
    rcvInfo.rcvnm = "수신자이름"
    
    '알림톡 내용, 최대 1000자
    rcvInfo.msg = "[테스트] 테스트 템플릿입니다."
    
    '대체문자 메시지 내용
    rcvInfo.altmsg = "대체문자 테스트입니다."
    
    rcvList.Add rcvInfo
    
    
    ReceiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, senderNum, "", "", altSendType, txtReserveDT.Text, rcvList)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    ' 알림톡 템플릿코드 - ListATStemplate API, GetURL(TEMPLATE) API, 또는 팝빌사이트에서 확인
    templateCode = "018020000001"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '알림톡 내용, 최대 1000자
    content = "[테스트] 테스트 템플릿입니다."
    
    '대체문자 내용
    altContent = "대체문자 테스트입니다."
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    ReceiptNum = KakaoService.SendATS(txtCorpNum.Text, templateCode, senderNum, content, altContent, altSendType, txtReserveDT.Text, rcvList)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    '첨부이미지 파일경로
    filePath = CommonDialog1.FileName
    
    '이미지 링크 URL
    imageURL = "http://www.popbill.com"
    
    '플러스친구 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '수신번호
    rcvInfo.rcv = "010000111"
    
    '수신자명
    rcvInfo.rcvnm = "수신자이름"
    
    '친구톡 내용, 최대 1000자
    rcvInfo.msg = "[테스트] 테스트 템플릿입니다."
    
    '대체문자 메시지 내용
    rcvInfo.altmsg = "대체문자 테스트입니다."
    
    rcvList.Add rcvInfo
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList, filePath, imageURL)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    '첨부이미지 파일경로
    filePath = CommonDialog1.FileName
    
    '이미지 링크 URL
    imageURL = "http://www.popbill.com"
    
    '플러스친구 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
                 
        '친구톡 내용, 최대 400자
        rcvInfo.msg = "친구톡 내용입니다. 수신자에 따라 다른 내용을 전송합니다." + CStr(i)
        
        '대체문자 메시지 내용
        rcvInfo.altmsg = "대체문자 내용입니다. 수신자에 따라 다른 내용을 전송할 수 있습니다." + CStr(i)
        
        rcvList.Add rcvInfo
    
    Next
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList, filePath, imageURL)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    '첨부이미지 파일경로
    filePath = CommonDialog1.FileName
    
    '이미지 링크 URL
    imageURL = "http://www.popbill.com"
    
    '플러스친구 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    '친구톡 내용, 최대 400자
    content = "친구톡 전송 내용입니다. 최대 1000자 까지 입력할 수 있습니다. 동일한 내용을 개별 수신자에게 전송하는 예제입니다."
    
    '대체문자 내용
    altContent = "대체문자 테스트입니다."
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFMS(txtCorpNum.Text, plusFriendID, senderNum, content, altContent, altSendType, txtReserveDT.Text, adsYN, rcvList, btnList, filePath, imageURL)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    '플러스친구 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
                 
        '친구톡 내용, 최대 1000자
        rcvInfo.msg = "친구톡 내용입니다. 수신자에 따라 다른 내용을 전송합니다." + CStr(i)
        
        '대체문자 메시지 내용
        rcvInfo.altmsg = "대체문자 내용입니다. 수신자에 따라 다른 내용을 전송할 수 있습니다." + CStr(i)
        
        rcvList.Add rcvInfo
    
    Next
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    
    '플러스친구 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    
    Set rcvInfo = New PBKakaoReceiver
    
    '수신번호
    rcvInfo.rcv = "010000111"
    
    '수신자명
    rcvInfo.rcvnm = "수신자이름"
    
    '친구톡 내용, 최대 1000자
    rcvInfo.msg = "[테스트] 테스트 템플릿입니다."
    
    '대체문자 메시지 내용
    rcvInfo.altmsg = "대체문자 테스트입니다."
    
    rcvList.Add rcvInfo
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, senderNum, "", "", altSendType, txtReserveDT.Text, adsYN, rcvList, btnList)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
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
    
    '플러스친구 아이디
    plusFriendID = "@팝빌"
    
    '팝빌에 사전 등록된 발신번호
    senderNum = "07043042993"
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
    
    '광고전송 여부
    adsYN = True
    
    '친구톡 내용, 최대 1000자
    content = "친구톡 전송 내용입니다. 최대 1000자 까지 입력할 수 있습니다. 동일한 내용을 개별 수신자에게 전송하는 예제입니다."
    
    '대체문자 내용
    altContent = "대체문자 테스트입니다."
    
    '대체문자 전송유형, 공백-미전송, C-알림톡내용 전송, A-대체문자내용 전송
    altSendType = ""
   
    '수신정보 배열, 최대 1000건
    For i = 0 To 99
    
        Set rcvInfo = New PBKakaoReceiver
        
        '수신번호
        rcvInfo.rcv = "0101122" + CStr(i)
        
        '수신자명
        rcvInfo.rcvnm = "수신자 이름" + CStr(i)
           
        rcvList.Add rcvInfo
    
    Next
    
    
    '버튼 정보구성, 최대 5개까지 배열에 추가 가능
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "버튼명"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbill.com"
    btnInfo.u2 = "http://wwww.linkhub.co.kr"
    
    btnList.Add btnInfo
    
    Set btnInfo = New PBKakaoButton
    
    btnInfo.n = "메시지전달"
    btnInfo.t = "MD"
    
    btnList.Add btnInfo
    
    
    ReceiptNum = KakaoService.SendFTS(txtCorpNum.Text, plusFriendID, senderNum, content, altContent, altSendType, txtReserveDT.Text, adsYN, rcvList, btnList)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

'=========================================================================
' 알림톡(ATS) 전송단가를 확인합니다.
'=========================================================================
Private Sub btnUnitCost_ATS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, ATS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "ATS 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 친구톡 이미지(FMS) 전송단가를 확인합니다.
'=========================================================================
Private Sub btnUnitCost_FMS_Click()
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "FMS 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 친구톡 텍스트(FTS) 전송단가를 확인합니다.
'=========================================================================
Private Sub btnUnitCost_FTS_Click(Index As Integer)
    Dim unitCost As Single
    
    unitCost = KakaoService.GetUnitCost(txtCorpNum.Text, FTS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "FTS 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-4304-2991"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '전체조회여부, Ture-회사조회, False-개인조회
    joinData.searchAllAllowYN = True
    
    Set Response = KakaoService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명
    CorpInfo.ceoname = "대표자"
    
    '상호
    CorpInfo.corpName = "상호"
    
    '주소
    CorpInfo.addr = "서울특별시"
    
    '업태
    CorpInfo.bizType = "업태"
    
    '종목
    CorpInfo.bizClass = "종목"
    
    Set Response = KakaoService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(KakaoService.LastErrCode) + vbCrLf + "응답메시지 : " + KakaoService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

Private Sub Form_Load()
    KakaoService.Initialize LinkID, SecretKey
    
    '연동환경 설정값 True-개발용, False-상업용
    KakaoService.IsTest = True
End Sub
