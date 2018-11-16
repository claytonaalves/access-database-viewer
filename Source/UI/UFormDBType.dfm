object FrmDBType: TFrmDBType
  Left = 800
  Top = 431
  BorderStyle = bsToolWindow
  Caption = 'Database Type'
  ClientHeight = 160
  ClientWidth = 179
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object LBDatabases: TListBox
    Left = 0
    Top = 0
    Width = 179
    Height = 160
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -21
    Font.Name = 'Tahoma'
    Font.Style = []
    ItemHeight = 25
    Items.Strings = (
      'MS Access'
      'Firebird'
      'MySQL'
      'SQLite')
    ParentFont = False
    TabOrder = 0
    OnDblClick = LBDatabasesDblClick
  end
end
