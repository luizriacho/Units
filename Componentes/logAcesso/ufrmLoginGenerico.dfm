object frmLoginGenerico: TfrmLoginGenerico
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Acesso ao Sistema'
  ClientHeight = 190
  ClientWidth = 350
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  KeyPreview = True
  Position = poScreenCenter
  OnKeyDown = FormKeyDown
  TextHeight = 15
  object lblUsuario: TLabel
    Left = 32
    Top = 24
    Width = 43
    Height = 15
    Caption = 'Usu'#225'rio:'
  end
  object lblSenha: TLabel
    Left = 32
    Top = 80
    Width = 35
    Height = 15
    Caption = 'Senha:'
  end
  object edtUsuario: TEdit
    Left = 32
    Top = 43
    Width = 285
    Height = 23
    TabOrder = 0
  end
  object edtSenha: TEdit
    Left = 32
    Top = 99
    Width = 285
    Height = 23
    PasswordChar = '*'
    TabOrder = 1
  end
  object btnEntrar: TBitBtn
    Left = 111
    Top = 140
    Width = 100
    Height = 32
    Caption = 'Entrar'
    Default = True
    NumGlyphs = 2
    TabOrder = 2
    OnClick = btnEntrarClick
  end
  object btnCancelar: TBitBtn
    Left = 217
    Top = 140
    Width = 100
    Height = 32
    Caption = 'Cancelar'
    Kind = bkCancel
    NumGlyphs = 2
    TabOrder = 3
    OnClick = btnCancelarClick
  end
end
