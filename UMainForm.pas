unit UMainForm;

interface

uses
  Windows, Classes, Forms, Messages, Controls, Dialogs, DBGrids, StdCtrls, Grids, Menus, ComCtrls, DB, SynEdit,
  SynHighlighterSQL, MSAccessU, UFileCatcher, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZDataset,
  ZAbstractDataset, SynEditHighlighter;

type
  TMainForm = class(TForm)
    ColumnPopupMenu: TPopupMenu;
    DBConnection: TZConnection;
    DataSource1: TDataSource;
    GridPopupMenu: TPopupMenu;
    LBTabelas: TListBox;
    MainGrid: TDBGrid;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    NmerodeRegistrosnatabela1: TMenuItem;
    OpenDialog1: TOpenDialog;
    Query: TZQuery;
    SqlEditor: TSynEdit;
    StatusBar: TStatusBar;
    SQLSyntaxHighliter: TSynSQLSyn;
    mnuAbrir: TMenuItem;
    mnuArquivo: TMenuItem;
    mnuColunas: TMenuItem;
    mnuEsconderColuna: TMenuItem;
    mnuFiltrar: TMenuItem;
    mnuGerarSelect: TMenuItem;
    mnuMedirDensidade: TMenuItem;
    mnuSair: TMenuItem;
    mnuTabelas: TMenuItem;
    mnuVisualizar: TMenuItem;
    procedure AbrirArquivoClick(Sender: TObject);
    procedure TabelaSelecionadaHandler(Sender: TObject);
    procedure SqlEditorKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure MainGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mnuSairClick(Sender: TObject);
    procedure mnuGerarSelectClick(Sender: TObject);
    procedure MainGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure mnuEsconderColunaClick(Sender: TObject);
    procedure MedirDensidadeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    FFileName: String;
    procedure PreencheListaDeTabelas;
    procedure AtualizaMenuPopupGrade;
    function LeArquivo(const FileName: String): MSAccessU.TByteArray;
    procedure AtivaDesativaColunaGrid(Sender: TObject);
    function FixQuotes(S: String): String;
    procedure WMDropFiles(var Msg: TWMDropFiles); message WM_DROPFILES;
    procedure OpenAccessDatabaseFile(const Filename: String);
    procedure OpenAccessDatabase;
    procedure OpenFirebirdDatabase;
    procedure OpenSQLiteDatabase;
    procedure OpenMySQLDatabase;
  end;

var
  MainForm: TMainForm;

implementation

uses SysUtils, ShellAPI, UFormDBType;

{$R *.dfm}

procedure TMainForm.FormCreate(Sender: TObject);
begin
   DragAcceptFiles(Handle, True);
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
   DragAcceptFiles(Handle, False);
end;

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
begin
   FrmDBType := TFrmDBType.Create(nil);

   if FrmDBType.ShowModal = mrOk then begin
      case FrmDBType.LBDatabases.ItemIndex of
         0: OpenAccessDatabase;
         1: OpenFirebirdDatabase;
         2: OpenMySQLDatabase;
         3: OpenSQLiteDatabase;
      end;
   end;
end;

procedure TMainForm.OpenMySQLDatabase;
begin
   DBConnection.Connected := False;
   DBConnection.Protocol := 'mysql-5';
   DBConnection.HostName := '10.1.1.100';
   DBConnection.User := 'root';
   DBConnection.Password := '1234';
   DBConnection.Database := 'vigo_erp';
   DBConnection.Connected := True;

   PreencheListaDeTabelas;
   Caption := '';
end;

procedure TMainForm.OpenSQLiteDatabase;
begin
   if not OpenDialog1.Execute then Exit;

   DBConnection.Connected := False;
   DBConnection.Protocol := 'sqlite-3';
   DBConnection.DataBase := OpenDialog1.FileName;
   DBConnection.Connected := True;

   PreencheListaDeTabelas;
   Caption := '';
end;

procedure TMainForm.OpenFirebirdDatabase;
begin
   if not OpenDialog1.Execute then Exit;

   DBConnection.Connected := False;
   DBConnection.Protocol := 'firebird-2.5';
   DBConnection.DataBase := OpenDialog1.FileName;
   DBConnection.Connected := True;

   PreencheListaDeTabelas;
   Caption := '';
end;

procedure TMainForm.OpenAccessDatabase;
begin
   if not OpenDialog1.Execute then Exit;
   OpenAccessDatabaseFile(OpenDialog1.FileName);
end;

procedure TMainForm.OpenAccessDatabaseFile(const Filename: String);
var Password: String;
begin
   Password := XorPassword(LeArquivo(FileName));

   DBConnection.Connected := False;
   DBConnection.Protocol := 'ado';
   DBConnection.DataBase :=
      'Provider=Microsoft.Jet.OLEDB.4.0; Data Source=' + FileName + ';Persist Security Info=False;Jet OLEDB:Database Password="'+ Password +'"';

   DBConnection.Connected := True;

   PreencheListaDeTabelas;
   Caption := FileName;
end;

procedure TMainForm.TabelaSelecionadaHandler(Sender: TObject);
var SelectedTableName: String;
begin
   SelectedTableName := LBTabelas.Items[LBTabelas.ItemIndex];

   Query.Active := False;
   Query.SQL.Text := 'SELECT * FROM ' + SelectedTableName;
   Query.Active :=  True;

   Caption := FFileName + ' - ' + SelectedTableName;

   StatusBar.Panels[1].Text := Format('%d registros', [Query.RecordCount]);
   StatusBar.Panels[2].Text := Format('%d colunas', [Query.FieldCount]);

   AtualizaMenuPopupGrade;
end;

procedure TMainForm.PreencheListaDeTabelas;
var TableNames: TStringList;
begin
   TableNames := TStringList.Create;
   DBConnection.GetTableNames('', TableNames);
   LBTabelas.Items := TableNames;
   TableNames.Free;
   StatusBar.Panels[0].Text := Format('%d tabelas', [LBTabelas.Count]);
end;

procedure TMainForm.SqlEditorKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if Key=116 then begin
      Query.SQL.Text := SqlEditor.Text;
      Query.Open;

      StatusBar.Panels[1].Text := Format('%d registros', [Query.RecordCount]);
      StatusBar.Panels[2].Text := Format('%d colunas', [Query.FieldCount]);
   end;
end;

procedure TMainForm.MainGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if (ssCtrl in Shift) and (Key=13) then begin
      SqlEditor.SetFocus;
   end;
end;

procedure TMainForm.mnuSairClick(Sender: TObject);
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

   TotalNumberRows := Query.RecordCount;
   Bookmark := Query.GetBookmark;

   Query.DisableControls;
   Query.First;

   sl := TStringList.Create;
   sl.Sorted := True;
   sl.Duplicates := dupIgnore;
   while not Query.Eof do begin
      sl.Add(MainGrid.SelectedField.AsString);
      Query.Next;
   end;
   DistinctValues := sl.Count;
   sl.Free;

   Selectivity := (1 / DistinctValues) * 100;

   ShowMessage(Format('%.4f%%', [Selectivity]));

   Query.GotoBookmark(Bookmark);
   Query.EnableControls;
end;

procedure TMainForm.WMDropFiles(var Msg: TWMDropFiles);
var
  I: Integer;                 // loops thru all dropped files
  DropPoint: TPoint;          // point where files dropped
  Catcher: TFileCatcher;      // file catcher class
begin
  Catcher := TFileCatcher.Create(Msg.Drop);
  try
    for I := 0 to Pred(Catcher.FileCount) do
    begin
      // ... code to process file here
      OpenAccessDatabaseFile(Catcher.Files[I]);
    end;
    DropPoint := Catcher.DropPoint;
    // ... do something with drop point
  finally
    Catcher.Free;
  end;
  Msg.Result := 0;
end;

end.

