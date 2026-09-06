object frmGestaoPerfil: TfrmGestaoPerfil
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Gest'#227'o de Perfis e Permiss'#245'es'
  ClientHeight = 480
  ClientWidth = 620
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnShow = FormShow
  TextHeight = 15
  object lblPerfis: TLabel
    Left = 16
    Top = 16
    Width = 100
    Height = 15
    Caption = 'Perfil Selecionado:'
  end
  object cbPerfis: TComboBox
    Left = 16
    Top = 37
    Width = 280
    Height = 23
    Style = csDropDownList
    TabOrder = 0
    OnChange = cbPerfisChange
  end
  object btnNovoPerfil: TBitBtn
    Left = 310
    Top = 36
    Width = 100
    Height = 25
    Caption = 'Novo Perfil'
    TabOrder = 1
    OnClick = btnNovoPerfilClick
  end
  object StringGridPerm: TStringGrid
    Left = 16
    Top = 75
    Width = 585
    Height = 340
    ColCount = 5
    DefaultRowHeight = 24
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    TabOrder = 2
  end
  object btnSalvar: TBitBtn
    Left = 405
    Top = 430
    Width = 95
    Height = 32
    Caption = 'Salvar'
    Kind = bkOK
    NumGlyphs = 2
    TabOrder = 3
    OnClick = btnSalvarClick
  end
  object btnFechar: TBitBtn
    Left = 506
    Top = 430
    Width = 95
    Height = 32
    Caption = 'Fechar'
    Kind = bkCancel
    NumGlyphs = 2
    TabOrder = 4
    OnClick = btnFecharClick
  end
end