unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, strutils, Vcl.ExtCtrls,
  Winapi.ActiveX, TlHelp32;

const
  SpotifyExecutable = '\Spotify.exe';

type
  TStatus = (stUnknown, stFound, stMuted);

  TMainForm = class(TForm)
    Timer: TTimer;
    procedure TimerTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    FStatus: TStatus;
    FSpotifyHandle: THandle;
    procedure Mute();
    procedure Unmute();
    { Private declarations }
  public
    { Public declarations }
    property Status: TStatus read FStatus write FStatus;
    property SpotifyHandle: THandle read FSpotifyHandle write FSpotifyHandle;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses uMutify;

var
  AudioEndpoints: IMMDeviceEnumerator;
  Speakers: IMMDevice;
  SessionManager: IAudioSessionManager2;
  GroupingString: string;

function isAdvertising(WindowTitle: String): Bool;
begin
  Result := not AnsiContainsText(WindowTitle, '-'); // Author - Song
end;

function ExeFromHandle(wHnd: THandle): String;
var
  Snap: THandle;
  Pe: tagPROCESSENTRY32;
  Pid: THandle;
  Found: Boolean;
begin
  Result := '';

  if (wHnd = 0) then
    Exit;

  Snap := TlHelp32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

  if (Snap = Cardinal(-1)) then
    Exit;

  GetWindowThreadProcessId(wHnd, @Pid);
  Pe.dwSize := Sizeof(Pe);
  Found := TlHelp32.Process32First(Snap, Pe);

  while Found do
  begin
    if (Pe.th32ProcessID = Pid) then
    begin
      Result := String(Pe.szExeFile);
      Break;
    end;
    Found := TlHelp32.Process32Next(Snap, Pe);
  end;

  CloseHandle(Snap);
end;

function isSpotify(wHnd: THandle): Bool;
begin
  Result := ExeFromHandle(wHnd) = 'Spotify.exe';
end;

function ProcessWindows(wHnd: THandle; Form: TMainForm): Bool; stdcall;
var
  WindowName, WindowTitle: array [0 .. 255] of char;
  msg: string;
  info: TWindowInfo;
begin
  Result := True;
  Form.SpotifyHandle := 0;

  if IsWindow(wHnd) then
  begin
    GetClassName(wHnd, WindowName, 255);
    if (AnsiStartsText('Chrome_Widget', WindowName)) then
    begin
      GetWindowText(wHnd, WindowTitle, 255);
      if (WindowTitle <> '') then
      begin
        if isSpotify(wHnd) and isAdvertising(WindowTitle) then
        begin
          Form.Status := stFound;
          Form.SpotifyHandle := wHnd;
          Result := False;
        end;
      end;
    end;
  end;
end;

procedure TMainForm.TimerTimer(Sender: TObject);
var
  CurrentVolume: Single;
  WindowTitle: array [0 .. 255] of char;
begin
  if IsWindow(FSpotifyHandle) then
  begin
    if (FStatus = stMuted) then
    begin
      GetWindowText(FSpotifyHandle, WindowTitle, 255);
      if not isAdvertising(WindowTitle) then
        Unmute;
    end
    else
    begin
      GetWindowText(FSpotifyHandle, WindowTitle, 255);
      if isAdvertising(WindowTitle) then
        Mute();
    end;
  end
  else
  begin
    EnumWindows(@ProcessWindows, LParam(Self));
    if (FStatus = stFound) then
      Mute();
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  HR: HResult;
begin
  HR := CoCreateInstance(CLSID_MMDeviceEnumerator, nil, CLSCTX_INPROC_SERVER,
    IID_IMMDeviceEnumerator, AudioEndpoints);
  HR := AudioEndpoints.GetDefaultAudioEndpoint(0, 1, Speakers);
  HR := Speakers.Activate(IID_IAudioSessionManager2, 0, nil, SessionManager);

  if HR <> S_OK then
  begin
    ShowMessage('Soemthing wrong here');
  end;

  Unmute();
end;

procedure TMainForm.Mute;
var
  SessionEnumerator: IAudioSessionEnumerator;
  SessionCount: integer;
  SessionControl: IAudioSessionControl;
  SessionControl2: IAudioSessionControl2;
  SimpleAudioVolume: ISimpleAudioVolume;
  SessionName: PWideChar;
  IconPath: PWideChar;
  SessionIdentifier: PWideChar;
  Grouping: TGUID;
  Context: TGUID;
  i: integer;
begin
  SessionManager.GetSessionEnumerator(SessionEnumerator);
  SessionEnumerator.GetCount(SessionCount);
  Context := GUID_NULL;

  for i := 0 to SessionCount - 1 do
  begin
    SessionEnumerator.GetSession(i, SessionControl);
    if SessionControl <> nil then
    begin
      SessionControl.GetIconPath(IconPath);
      SessionControl.GetDisplayName(SessionName);
      SessionControl.GetGroupingParam(Grouping);

      SessionControl.QueryInterface(IID_IAudioSessionControl2, SessionControl2);
      SessionControl2.GetSessionInstanceIdentifier(SessionIdentifier);

      if Pos(SpotifyExecutable, SessionIdentifier) > 0 then
      begin
        SessionControl.QueryInterface(IID_ISimpleAudioVolume,
          SimpleAudioVolume);
        if SimpleAudioVolume <> nil then
        begin
          SimpleAudioVolume.SetMute(true, Context);
          SimpleAudioVolume := nil;
          FStatus := stMuted;
        end;
      end;

      CoTaskMemFree(IconPath);
      CoTaskMemFree(SessionName);
      CoTaskMemFree(SessionIdentifier);
      SessionControl2 := nil;
      SessionControl := nil;
    end;
  end;
  SessionEnumerator := nil;
end;

procedure TMainForm.Unmute;
var
  SessionEnumerator: IAudioSessionEnumerator;
  SessionCount: integer;
  SessionControl: IAudioSessionControl;
  SessionControl2: IAudioSessionControl2;
  SimpleAudioVolume: ISimpleAudioVolume;
  SessionName: PWideChar;
  IconPath: PWideChar;
  SessionIdentifier: PWideChar;
  Grouping: TGUID;
  Context: TGUID;
  i: integer;
begin
  SessionManager.GetSessionEnumerator(SessionEnumerator);
  SessionEnumerator.GetCount(SessionCount);
  Context := GUID_NULL;

  for i := 0 to SessionCount - 1 do
  begin
    SessionEnumerator.GetSession(i, SessionControl);
    if SessionControl <> nil then
    begin
      SessionControl.GetIconPath(IconPath);
      SessionControl.GetDisplayName(SessionName);
      SessionControl.GetGroupingParam(Grouping);

      SessionControl.QueryInterface(IID_IAudioSessionControl2, SessionControl2);
      SessionControl2.GetSessionInstanceIdentifier(SessionIdentifier);

      if Pos(SpotifyExecutable, SessionIdentifier) > 0 then
      begin
        SessionControl.QueryInterface(IID_ISimpleAudioVolume,
          SimpleAudioVolume);
        if SimpleAudioVolume <> nil then
        begin
          SimpleAudioVolume.SetMute(False, Context);
          SimpleAudioVolume := nil;
          FStatus := stUnknown;
        end;
      end;

      CoTaskMemFree(IconPath);
      CoTaskMemFree(SessionName);
      CoTaskMemFree(SessionIdentifier);
      SessionControl2 := nil;
      SessionControl := nil;
    end;
  end;
  SessionEnumerator := nil;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  SessionManager := nil;
  Speakers := nil;
  AudioEndpoints := nil;
end;

end.
