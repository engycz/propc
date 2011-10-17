program prDemo7;
{$R 'prDemo7.res' 'prDemo7.rc'}

uses
  Messages,
  SysUtils,
  Windows,
  OpcServerUnit in 'OpcServerUnit.pas';

const
  APPNAME = 'PRDEMO7';

{$I prDemo7.inc} {Resource constants}

function WindowProc(hWindow: HWND; iMsg: UINT; wPar: WPARAM; lPar: LPARAM): LRESULT; stdcall;
begin
  Result:= 0;
  case iMsg of
    WM_COMMAND:
    case LOWORD(wPar) of
      IDCANCEL:	SendMessage(hWindow, WM_CLOSE, 0, 0);
    end;
    WM_CLOSE:
    if CallTerminateProcs then
      DestroyWindow(hWindow);
    WM_DESTROY:
    begin
      PostQuitMessage(0);
    end
  else
    Result:= DefWindowProc(hWindow, iMsg, wPar, lPar)
  end
end;

procedure WinMain;
var
  hWindow: HWND;
  msg: TMsg;
  wc: TWndClassEx;
begin
  with wc do
  begin
    cbSize:= sizeof(wc);
    style:= CS_HREDRAW or CS_VREDRAW;
    lpfnWndProc:= @WindowProc;
    cbClsExtra:= 0;
    cbWndExtra:= DLGWINDOWEXTRA;
    hInstance:= System.MainInstance;
    hIcon:= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MAIN));
    hCursor:= LoadCursor(0, IDC_ARROW);
    hbrBackground:= COLOR_WINDOW;
    lpszMenuName:= nil;
    lpszClassName:= APPNAME;
    hIconSm:= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MAIN));
  end;
  RegisterClassEx(wc);
  hWindow:= CreateDialog( hInstance, MAKEINTRESOURCE(IDD_MAIN), 0, nil);
  MsgWnd:= GetDlgItem(hWindow, IDC_CLIENTMSG);
  ShowWindow(hWindow, SW_SHOWNORMAL);
  while GetMessage (msg, 0, 0, 0) do
  if not IsDialogMessage(hWindow, msg) then
  begin
    TranslateMessage(msg);
    DispatchMessage(msg)
  end
end;

type
  TProcedure = procedure;
begin
  if Assigned(InitProc) then
     TProcedure(InitProc);
  WinMain
end.