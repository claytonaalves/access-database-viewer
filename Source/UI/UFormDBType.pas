unit UFormDBType;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TFrmDBType = class(TForm)
    LBDatabases: TListBox;
    procedure LBDatabasesDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
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

procedure TFrmDBType.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   case Key of
      13: ModalResult := mrOk;
      27: ModalResult := mrCancel;
   end;
end;

end.
