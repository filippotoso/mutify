program Mutify;

uses
  Vcl.Forms,
  uMainForm in 'uMainForm.pas' {MainForm},
  uMutify in 'uMutify.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.ShowMainForm := False;
  Application.MainFormOnTaskbar := False;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;

end.
