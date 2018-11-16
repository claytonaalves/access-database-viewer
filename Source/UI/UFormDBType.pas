unit UFormDBType;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TFrmDBType = class(TForm)
    LBDatabases: TListBox;
    procedure LBDatabasesDblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmDBType: TFrmDBType;

implementation

{$R *.dfm}

procedure TFrmDBType.LBDatabasesDblClick(Sender: TObject);
begin
   ModalResult := mrOk;
end;

end.
