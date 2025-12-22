#include <vcl.h>
#pragma hdrstop

#include "ServiceController.h"
#include <utilcls.h> // TNoParam, TVariant용

#include <windows.h>
#include <stdio.h>

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"

TGabbianiAgent *GabbianiAgent;

//---------------------------------------------------------------------------
__fastcall TGabbianiAgent::TGabbianiAgent(TComponent* Owner)
    : TService(Owner)
{
    StartFlag = false;
}

//---------------------------------------------------------------------------
TServiceController __fastcall TGabbianiAgent::GetServiceController(void)
{
    return (TServiceController) ServiceController;
}

// ---------------------------------------------------------------------------
// Win32 API를 직접 사용하여 로그를 남기는 함수
// ---------------------------------------------------------------------------
void WriteLogAPI(AnsiString msg)
{
    HANDLE hFile;
    DWORD dwBytesWritten;

    // 1. 파일 경로 설정 (C:\Logs 폴더가 없으면 에러날 수 있으므로 C:\ 루트나 Temp 권장)
    // 주의: 백슬래시 2개(\\) 필수
    AnsiString FileName = "E:\\GabbianiAgent\\Gabbiani_Service_Log.txt";

    // 2. 파일 열기 또는 생성 (API 함수)
    hFile = CreateFile(
        FileName.c_str(),           // 파일 경로
        GENERIC_WRITE,              // 쓰기 권한
        FILE_SHARE_READ,            // 다른 프로그램이 읽을 수 있게 함
        NULL,                       // 보안 속성 (기본값)
        OPEN_ALWAYS,                // 파일이 있으면 열고, 없으면 새로 만듦 (중요!)
        FILE_ATTRIBUTE_NORMAL,      // 일반 파일 속성
        NULL
    );

    if (hFile != INVALID_HANDLE_VALUE)
    {
        // 3. 파일 포인터를 맨 끝으로 이동 (Append 모드 구현)
        SetFilePointer(hFile, 0, NULL, FILE_END);

        // 4. 시간 문자열 생성
        SYSTEMTIME st;
        GetLocalTime(&st);

        char timeBuf[100];
        sprintf(timeBuf, "[%02d:%02d:%02d] ", st.wHour, st.wMinute, st.wSecond);

        // 5. 시간 기록
        WriteFile(hFile, timeBuf, strlen(timeBuf), &dwBytesWritten, NULL);

        // 6. 메시지 + 줄바꿈 기록
        AnsiString finalMsg = msg + "\r\n"; // 윈도우 줄바꿈(\r\n)
        WriteFile(hFile, finalMsg.c_str(), finalMsg.Length(), &dwBytesWritten, NULL);

        // 7. 핸들 닫기 (메모리 누수 방지)
        CloseHandle(hFile);
    }
}

// ---------------------------------------------------------------------------
// [1] 메시지를 타임스탬프와 함께 메모리(m_LogBuffer)에만 저장하는 함수
// ---------------------------------------------------------------------------
void __fastcall TGabbianiAgent::AddLogToBuffer(AnsiString msg)
{
    SYSTEMTIME st;
    GetLocalTime(&st);
    char timeBuf[100];
    
    // 발생한 시점의 시간을 기록합니다.
    sprintf(timeBuf, "[%02d:%02d:%02d] ", st.wHour, st.wMinute, st.wSecond);
    
    // 버퍼에 "시간 + 메시지 + 줄바꿈"을 누적합니다.
    m_LogBuffer += (AnsiString(timeBuf) + msg + "\r\n");
}

// ---------------------------------------------------------------------------
// [2] 모아둔 버퍼 내용을 파일로 내보내는 함수 (기존 Win32 API 코드 활용)
// ---------------------------------------------------------------------------
void __fastcall TGabbianiAgent::FlushLogToFile()
{
    if (m_LogBuffer.IsEmpty()) return; // 쓸 내용 없으면 리턴

    HANDLE hFile;
    DWORD dwBytesWritten;
    AnsiString FileName = "E:\\GabbianiAgent\\Gabbiani_Service_Log.txt";

    hFile = CreateFile(
        FileName.c_str(), GENERIC_WRITE, FILE_SHARE_READ, NULL,
        OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL
    );

    if (hFile != INVALID_HANDLE_VALUE)
    {
        SetFilePointer(hFile, 0, NULL, FILE_END);
        
        // 모아둔 전체 문자열(m_LogBuffer)을 한 번에 기록
        WriteFile(hFile, m_LogBuffer.c_str(), m_LogBuffer.Length(), &dwBytesWritten, NULL);
        
        CloseHandle(hFile);
    }

    // 기록 후 버퍼 초기화 (중복 기록 방지)
    m_LogBuffer = ""; 
}

//
void __stdcall ServiceController(unsigned CtrlCode)
{
    GabbianiAgent->Controller(CtrlCode);
}

//---------------------------------------------------------------------------

void __fastcall TGabbianiAgent::ServiceStart(TService *Sender, bool &Started)
{

    try
    {

		AddLogToBuffer("1. Service Start 진입");

        // 1. TMS Async32 설정
        MyComm = new TVaComm(NULL);
        MyComm->PortNum = 3;          // ★ 포트 번호 확인 (장치관리자)
        MyComm->Baudrate = br9600;
        MyComm->Open();

        // 2. OPC 서버 연결
        OPCServer = CoOPCServer::Create();
        OPCServer->Connect(WideString("Schneider-Aut.OFS"), TNoParam());

    	// [2단계] 설정 로딩
    	AddLogToBuffer("2. 초기화 진행 중...");

        // 3. 그룹 추가
        OPCGroups = OPCServer->OPCGroups;

        // (그룹은 'I'가 붙어 있습니다)
        IOPCGroup *tempGroup = NULL;
        OPCGroups->Add(TVariant(WideString("GabbianiGroup")), &tempGroup);
        MyGroup = tempGroup;

        // 4. 아이템 등록
        MyItems = MyGroup->OPCItems;

        // ★★★ [수정됨: I가 없는 OPCItem 사용] ★★★
        OPCItem *tempItem1 = NULL;
        OPCItem *tempItem2 = NULL;

        // 주소 등록 (예시 주소)
        MyItems->AddItem(WideString("SS!%MW1470"), 1, &tempItem1);
        MyItems->AddItem(WideString("SS!%MW1461"), 2, &tempItem2);

        ItemStatus = tempItem1;
        ItemSpeed  = tempItem2;

        // 5. 타이머 설정
        if (Timer1 == NULL) {
            Timer1 = new TTimer(NULL);
            Timer1->OnTimer = Timer1Timer;
        }
        Timer1->Interval = 1000;
        Timer1->Enabled = true;

        Started = true;

    	// [3단계] 네트워크/리소스 준비
    	// ... 준비 코드 ...
    	AddLogToBuffer("3. 서비스 시작 성공!");

    }
    catch (...)
    {
        Started = false;
        WriteLogAPI("FATAL: 알 수 없는 오류");
    }
}

//---------------------------------------------------------------------------
void __fastcall TGabbianiAgent::ServiceStop(TService *Sender, bool &Stopped)
{
    Timer1->Enabled = false;

    try {
        if (OPCServer) OPCServer->Disconnect();
    } catch (...) {}

    if (MyComm) {
        MyComm->Close();
        delete MyComm;
        MyComm = NULL;
    }
    Stopped = true;
}

//---------------------------------------------------------------------------
void __fastcall TGabbianiAgent::Timer1Timer(TObject *Sender)
{
    Timer1->Enabled = false;

	// 1. ServiceStart에서 모아둔 로그가 있다면 파일에 기록하고 비움
    if (!m_LogBuffer.IsEmpty())
    {
        FlushLogToFile();
    }

    try
    {
        TVariant valStatus, valSpeed;
        
        // 데이터 읽기
        if (ItemStatus) ItemStatus->get_Value(&valStatus);
        if (ItemSpeed)  ItemSpeed->get_Value(&valSpeed);

        int iStatus = int(valStatus);
        int iSpeed  = int(valSpeed);

        // 데이터 전송
        String sendData;
        sendData.printf("$ST,%d,SP,%d\r\n", iStatus, iSpeed);

        if (MyComm) {
             MyComm->WriteText(sendData);
        }
    }
    catch (...) {
        // 무시
    }

    Timer1->Enabled = true;
}
//---------------------------------------------------------------------------
