object frmCadastroUsuario: TfrmCadastroUsuario
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Cadastro de Usu'#225'rios'
  ClientHeight = 330
  ClientWidth = 420
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnShow = FormShow
  TextHeight = 15
  object lblNome: TLabel
    Left = 24
    Top = 20
    Width = 87
    Height = 15
    Caption = 'Nome Completo:'
  end
  object lblPerfil: TLabel
    Left = 24
    Top = 75
    Width = 84
    Height = 15
    Caption = 'Perfil de Acesso:'
  end
  object lblLogin: TLabel
    Left = 24
    Top = 130
    Width = 33
    Height = 15
    Caption = 'Login:'
  end
  object lblSenha: TLabel
    Left = 210
    Top = 130
    Width = 35
    Height = 15
    Caption = 'Senha:'
  end
  object edtNome: TEdit
    Left = 24
    Top = 40
    Width = 370
    Height = 23
    TabOrder = 0
  end
  object cbPerfil: TComboBox
    Left = 24
    Top = 95
    Width = 370
    Height = 23
    Style = csDropDownList
    TabOrder = 1
  end
  object edtLogin: TEdit
    Left = 24
    Top = 150
    Width = 170
    Height = 23
    TabOrder = 2
  end
  object edtSenha: TEdit
    Left = 210
    Top = 150
    Width = 184
    Height = 23
    PasswordChar = '*'
    TabOrder = 3
  end
  object chkAtivo: TCheckBox
    Left = 24
    Top = 190
    Width = 97
    Height = 17
    Caption = 'Usu'#225'rio Ativo'
    Checked = True
    State = cbChecked
    TabOrder = 4
  end
  object btnSalvar: TBitBtn
    Left = 204
    Top = 275
    Width = 90
    Height = 32
    Caption = 'Salvar'
    Kind = bkOK
    NumGlyphs = 2
    TabOrder = 5
    OnClick = btnSalvarClick
  end
  object btnCancelar: TBitBtn
    Left = 304
    Top = 275
    Width = 90
    Height = 32
    Caption = 'Cancelar'
    Kind = bkCancel
    NumGlyphs = 2
    TabOrder = 6
    OnClick = btnCancelarClick
  end
end