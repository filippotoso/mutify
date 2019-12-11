unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, strutils, Vcl.ExtCtrls;

type
  TStatus = (stUnknown, stFound, stMuted);

  TMainForm = class(TForm)
    Timer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
  private
    FStatus: TStatus;
    FVolume: Single;
    { Private declarations }
  public
    { Public declarations }
    property Status: TStatus read FStatus write FStatus;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses uMutify;

function ProcessWindows(wHnd: THandle; Form: TMainForm): Bool; stdcall;
var
  WindowName, WindowTitle: array [0 .. 255] of char;
  msg: string;
  info: TWindowInfo;
begin
  Result := True;

  if IsWindow(wHnd) then
  begin
    GetClassName(wHnd, WindowName, 255);
    if (WindowName = 'Chrome_WidgetWin_0') then
    begin
      GetWindowText(wHnd, WindowTitle, 255);
      if (WindowTitle <> '') then
      begin
        if (WindowTitle = 'Advertisement') then
          Form.Status := stFound;
        Result := False;
      end;
    end;
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  FVolume := GetMasterVolume();
end;

procedure TMainForm.TimerTimer(Sender: TObject);
begin
  if (FStatus <> stMuted) then
    FStatus := stUnknown;

  EnumWindows(@ProcessWindows, LParam(Self));
  if (FStatus = stFound) then
  begin
    FVolume := GetMasterVolume();
    SetMasterVolume(0);
    FStatus := stMuted;
  end
  else if (FStatus = stMuted) then
    SetMasterVolume(FVolume);
end;

end.
