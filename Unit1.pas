unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, DB, ADODB, StdCtrls, SynEdit,
  SynEditHighlighter, SynHighlighterSQL;

type
  TForm1 = class(TForm)
    ADOConnection1: TADOConnection;
    ADOTable1: TADOTable;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    LBTabelas: TListBox;
    ADOQuery1: TADOQuery;
    SynEdit1: TSynEdit;
    SynSQLSyn1: TSynSQLSyn;
    procedure Button1Click(Sender: TObject);
    procedure LBTabelasClick(Sender: TObject);
    procedure SynEdit1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    FFileName: String;
    procedure PreencheListaDeTabelas;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
   if not OpenDialog1.Execute then Exit;

   FFileName := OpenDialog1.FileName;

   ADOConnection1.Connected := False;
   ADOConnection1.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0;' +
      'User ID=Admin;' +
      'Data Source=' + FFileName + ';' +
      'Mode=Share Deny None;' +
      'Jet OLEDB:System database="";' +
      'Jet OLEDB:Registry Path="";' +
      'Jet OLEDB:Database Password="";' +
      'Jet OLEDB:Engine Type=5;' +
      'Jet OLEDB:Database Locking Mode=1;' +
      'Jet OLEDB:Global Partial Bulk Ops=2;' +
      'Jet OLEDB:Global Bulk Transactions=1;' +
      'Jet OLEDB:New Database Password="";' +
      'Jet OLEDB:Create System Database=False;' +
      'Jet OLEDB:Encrypt Database=False;' +
      'Jet OLEDB:Don''t Copy Locale on Compact=False;' +
      'Jet OLEDB:Compact Without Replica Repair=False;' +
      'Jet OLEDB:SFP=False;';
   ADOConnection1.Connected := True;
   PreencheListaDeTabelas;
   Caption := FFileName;
end;

procedure TForm1.LBTabelasClick(Sender: TObject);
var SelectedTableName: String;
begin
   DataSource1.DataSet := ADOTable1;

   SelectedTableName := LBTabelas.Items[LBTabelas.ItemIndex];

   ADOTable1.Active := False;
   ADOTable1.TableName := SelectedTableName;
   ADOTable1.Active :=  True;

   Caption := FFileName + ' - ' + SelectedTableName;
end;

procedure TForm1.PreencheListaDeTabelas;
var TableNames: TStringList;
begin
   TableNames := TStringList.Create;
   ADOConnection1.GetTableNames(TableNames);

   LBTabelas.Items := TableNames;

   TableNames.Free;
end;

procedure TForm1.SynEdit1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if Key=116 then begin
      DataSource1.DataSet := ADOQuery1;
      ADOQuery1.SQL.Text := SynEdit1.Text;
      ADOQuery1.Open;
   end;
end;

procedure TForm1.DBGrid1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if (ssCtrl in Shift) and (Key=13) then begin
      SynEdit1.SetFocus;
   end;
end;

end.
