program AccessDatabaseViewer;

uses
  Forms,
  UMainForm in 'UMainForm.pas' {MainForm},
  MSAccessU in 'Source\MSAccessU.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
