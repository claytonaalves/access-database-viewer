program AccessDatabaseViewer;

uses
  Forms,
  UMainForm in 'UMainForm.pas' {MainForm},
  MSAccessU in 'Source\MSAccessU.pas',
  UFileCatcher in 'Source\UFileCatcher.pas',
  UFormDBType in 'Source\UI\UFormDBType.pas' {FrmDBType};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
