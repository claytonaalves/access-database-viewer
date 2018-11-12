unit UMainForm;

interface

uses
  Windows, Classes, Forms, Controls, Dialogs, DBGrids, StdCtrls, Grids, DB, SynEdit, ADODB, SynHighlighterSQL, SynEditHighlighter, MSAccessU,
  Menus,
  ComCtrls;

type
  TMainForm = class(TForm)
    ADOConnection1: TADOConnection;
    ADOTable1: TADOTable;
    DataSource1: TDataSource;
    OpenDialog1: TOpenDialog;
    LBTabelas: TListBox;
    ADOQuery1: TADOQuery;
    SqlEditor: TSynEdit;
    SynSQLSyn1: TSynSQLSyn;
    MainMenu1: TMainMenu;
    Arquivo1: TMenuItem;
    Visualizar1: TMenuItem;
    Abrir1: TMenuItem;
    N1: TMenuItem;
    Sair1: TMenuItem;
    StatusBar: TStatusBar;
    abelas1: TMenuItem;
    Filtrar1: TMenuItem;
    NmerodeRegistrosnatabela1: TMenuItem;
    GridPopupMenu: TPopupMenu;
    mnuColunas: TMenuItem;
    MainGrid: TDBGrid;
    N2: TMenuItem;
    mnuGerarSelect: TMenuItem;
    ColumnPopupMenu: TPopupMenu;
    mnuEsconderColuna: TMenuItem;
    mnuMedirDensidade: TMenuItem;
    procedure AbrirArquivoClick(Sender: TObject);
    procedure TabelaSelecionadaHandler(Sender: TObject);
    procedure SqlEditorKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure MainGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Sair1Click(Sender: TObject);
    procedure mnuGerarSelectClick(Sender: TObject);
    procedure MainGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure mnuEsconderColunaClick(Sender: TObject);
    procedure MedirDensidadeClick(Sender: TObject);
  private
    FFileName: String;
    procedure PreencheListaDeTabelas;
    procedure AtualizaMenuPopupGrade;
    function LeArquivo(const FileName: String): MSAccessU.TByteArray;
    procedure AtivaDesativaColunaGrid(Sender: TObject);
    function FixQuotes(S: String): String;
  end;

var
  MainForm: TMainForm;

implementation

uses SysUtils;

{$R *.dfm}

function TMainForm.LeArquivo(const FileName: String): MSAccessU.TByteArray;
var Arq: File;
begin
   AssignFile(Arq, FileName);
   FileMode := fmOpenRead;
   Reset(Arq, 1);
   BlockRead(Arq, Result, 85);
   CloseFile(Arq);
end;

procedure TMainForm.AbrirArquivoClick(Sender: TObject);
var Password: String;
begin
   if not OpenDialog1.Execute then Exit;

   FFileName := OpenDialog1.FileName;

   Password := XorPassword(LeArquivo(FFileName));

   ADOConnection1.Connected := False;
   ADOConnection1.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0; Data Source=' + FFileName + ';Persist Security Info=False;Jet OLEDB:Database Password="'+ Password +'"';

   ADOConnection1.Connected := True;
   PreencheListaDeTabelas;
   Caption := FFileName;
end;

procedure TMainForm.TabelaSelecionadaHandler(Sender: TObject);
var SelectedTableName: String;
begin
   DataSource1.DataSet := ADOTable1;

   SelectedTableName := LBTabelas.Items[LBTabelas.ItemIndex];

   ADOTable1.Active := False;
   ADOTable1.TableName := SelectedTableName;
   ADOTable1.Active :=  True;

   Caption := FFileName + ' - ' + SelectedTableName;

   StatusBar.Panels[1].Text := Format('%d registros', [ADOTable1.RecordCount]);
   StatusBar.Panels[2].Text := Format('%d colunas', [ADOTable1.FieldCount]);

   AtualizaMenuPopupGrade;
end;

procedure TMainForm.PreencheListaDeTabelas;
var TableNames: TStringList;
begin
   TableNames := TStringList.Create;
   ADOConnection1.GetTableNames(TableNames);
   LBTabelas.Items := TableNames;
   TableNames.Free;
   StatusBar.Panels[0].Text := Format('%d tabelas', [LBTabelas.Count]);
end;

procedure TMainForm.SqlEditorKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if Key=116 then begin
      DataSource1.DataSet := ADOQuery1;
      ADOQuery1.SQL.Text := SqlEditor.Text;
      ADOQuery1.Open;

      StatusBar.Panels[1].Text := Format('%d registros', [ADOQuery1.RecordCount]);
      StatusBar.Panels[2].Text := Format('%d colunas', [ADOQuery1.FieldCount]);
   end;
end;

procedure TMainForm.MainGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if (ssCtrl in Shift) and (Key=13) then begin
      SqlEditor.SetFocus;
   end;
end;

procedure TMainForm.Sair1Click(Sender: TObject);
begin
   Application.Terminate;
end;

procedure TMainForm.AtualizaMenuPopupGrade;
var
   i: Integer;
   MenuItem: TMenuItem;
begin
   GridPopupMenu.Items[0].Clear;
   for i := 0 to MainGrid.Columns.Count - 1 do begin
      MenuItem := TMenuItem.Create(GridPopupMenu);
      MenuItem.Caption := MainGrid.Columns[i].Title.DefaultCaption;
      MenuItem.Checked := True;
      MenuItem.OnClick := AtivaDesativaColunaGrid;
      GridPopupMenu.Items[0].Add(MenuItem);
   end;
end;

procedure TMainForm.AtivaDesativaColunaGrid(Sender: TObject);
var MenuItem: TMenuItem;
begin
   MenuItem := TMenuItem(Sender);
   MenuItem.Checked := not MenuItem.Checked;
   MainGrid.Columns[MenuItem.MenuIndex].Visible := MenuItem.Checked;
end;

procedure TMainForm.mnuGerarSelectClick(Sender: TObject);
var
   i: Integer;
   Column: TColumn;
   ColumnNames: String;
   SelectedColumns: TStringList;
begin
   SelectedColumns := TStringList.Create;
   SelectedColumns.Delimiter := ',';

   for i := 0 to MainGrid.Columns.Count - 1 do begin
      Column := MainGrid.Columns[i];
      if Column.Visible then
         SelectedColumns.Add(Column.Field.FieldName);
   end;

   ColumnNames := FixQuotes(SelectedColumns.DelimitedText);

   SqlEditor.Text := 'SELECT ' + StringReplace(ColumnNames, ',', ', ', [rfReplaceAll]) + ' FROM ?';

   SelectedColumns.Free;
end;

function TMainForm.FixQuotes(S: String): String;
var
   OpenQuote: Boolean;
   i: Integer;
   C: Char;
begin
   Result := '';
   OpenQuote := False;
   for i := 1 to Length(S) do begin
      C := S[i];
      if C = '"' then begin
         if OpenQuote
         then C := ']'
         else C := '[';
         OpenQuote := not OpenQuote;
      end;
      Result := Result + C;
   end;
end;

procedure TMainForm.MainGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
   Coord: TGridCoord;
   Column: TColumn;
begin
   if Button = mbRight then begin
      Coord := MainGrid.MouseCoord(X, Y);
      if Coord.X = -1 then Exit;
      Column := MainGrid.Columns[Coord.X - 1];

      mnuEsconderColuna.Tag := Integer(Column);
      mnuMedirDensidade.Tag := Column.Index;

      ColumnPopupMenu.Popup(Self.Left + MainGrid.Left + X, Self.Top + MainGrid.Top + Y + 60);
   end;
end;

procedure TMainForm.mnuEsconderColunaClick(Sender: TObject);
var Column: TColumn;
begin
   Column := TColumn(mnuEsconderColuna.Tag);
   Column.Visible := False;

   StatusBar.Panels[2].Text := Format('%d colunas', [MainGrid.Columns.Count]);
end;

procedure TMainForm.MedirDensidadeClick(Sender: TObject);
var
   sl: TStringList;
   Bookmark: Pointer;
   TotalNumberRows, DistinctValues: Integer;
   Selectivity: Double;
begin
   MainGrid.SelectedIndex := mnuMedirDensidade.Tag;

   TotalNumberRows := ADOQuery1.RecordCount;
   Bookmark := ADOQuery1.GetBookmark;

   ADOQuery1.DisableControls;
   ADOQuery1.First;

   sl := TStringList.Create;
   sl.Sorted := True;
   sl.Duplicates := dupIgnore;
   while not ADOQuery1.Eof do begin
      sl.Add(MainGrid.SelectedField.AsString);
      ADOQuery1.Next;
   end;
   DistinctValues := sl.Count;
   sl.Free;

   Selectivity := (1 / DistinctValues) * 100;

   ShowMessage(Format('%.4f%%', [Selectivity]));

   ADOQuery1.GotoBookmark(Bookmark);
   ADOQuery1.EnableControls;
end;

end.

