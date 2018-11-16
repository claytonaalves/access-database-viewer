program DatabaseViewer;

uses
  Forms,
  MSAccessU in 'Source\MSAccessU.pas',
  UFileCatcher in 'Source\UFileCatcher.pas',
  UMainForm in 'Source\UI\UMainForm.pas' {MainForm},
  UFormDBType in 'Source\UI\UFormDBType.pas' {FrmDBType};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
