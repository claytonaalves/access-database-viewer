program AccessDatabaseViewer;

uses
  Forms,
  UMainForm in 'UMainForm.pas' {MainForm},
  MSAccessU in 'Source\MSAccessU.pas',
  UFileCatcher in 'Source\UFileCatcher.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
