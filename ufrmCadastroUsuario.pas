unit ufrmCadastroUsuario;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, 
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  FireDAC.Comp.Client, System.Hash;

type
  TfrmCadastroUsuario = class(TForm)
    edtNome: TEdit;
    cbPerfil: TComboBox;
    edtLogin: TEdit;
    edtSenha: TEdit;
    chkAtivo: TCheckBox;
    lblNome: TLabel;
    lblPerfil: TLabel;
    lblLogin: TLabel;
    lblSenha: TLabel;
    btnSalvar: TBitBtn;
    btnCancelar: TBitBtn;
    procedure btnSalvarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FConnection: TFDConnection;
    FIdUsuario: Integer;
    procedure CarregarPerfis;
  public
    class procedure NovoUsuario(AConnection: TFDConnection);
    class procedure EditarUsuario(AConnection: TFDConnection; AIdUsuario: Integer);
  end;

implementation

{$R *.dfm}

class procedure TfrmCadastroUsuario.NovoUsuario(AConnection: TFDConnection);
var
  Frm: TfrmCadastroUsuario;
begin
  Frm := TfrmCadastroUsuario.Create(nil);
  try
    Frm.FConnection := AConnection;
    Frm.FIdUsuario  := 0;
    Frm.ShowModal;
  finally
    Frm.Free;
  end;
end;

class procedure TfrmCadastroUsuario.EditarUsuario(AConnection: TFDConnection; AIdUsuario: Integer);
var
  Frm: TfrmCadastroUsuario;
begin
  Frm := TfrmCadastroUsuario.Create(nil);
  try
    Frm.FConnection := AConnection;
    Frm.FIdUsuario  := AIdUsuario;
    Frm.ShowModal;
  finally
    Frm.Free;
  end;
end;

procedure TfrmCadastroUsuario.CarregarPerfis;
var
  Qry: TFDQuery;
begin
  cbPerfil.Items.Clear;
  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := FConnection;
    Qry.SQL.Text := 'SELECT ID_PERFIL, NOME_PERFIL FROM PERFIL ORDER BY NOME_PERFIL';
    Qry.Open;

    while not Qry.Eof do
    begin
      cbPerfil.Items.AddObject(Qry.FieldByName('NOME_PERFIL').AsString, TObject(Qry.FieldByName('ID_PERFIL').AsInteger));
      Qry.Next;
    end;

    if cbPerfil.Items.Count > 0 then
      cbPerfil.ItemIndex := 0;
  finally
    Qry.Free;
  end;
end;

procedure TfrmCadastroUsuario.FormShow(Sender: TObject);
var
  Qry: TFDQuery;
  I, IdPerfilUser: Integer;
begin
  CarregarPerfis;

  if FIdUsuario > 0 then
  begin
    Qry := TFDQuery.Create(nil);
    try
      Qry.Connection := FConnection;
      Qry.SQL.Text := 'SELECT ID_PERFIL, NOME, LOGIN, ATIVO FROM USUARIO WHERE ID_USUARIO = :ID';
      Qry.ParamByName('ID').AsInteger := FIdUsuario;
      Qry.Open;

      if not Qry.IsEmpty then
      begin
        edtNome.Text  := Qry.FieldByName('NOME').AsString;
        edtLogin.Text := Qry.FieldByName('LOGIN').AsString;
        chkAtivo.Checked := Qry.FieldByName('ATIVO').AsString = 'S';
        
        IdPerfilUser := Qry.FieldByName('ID_PERFIL').AsInteger;
        for I := 0 to cbPerfil.Items.Count - 1 do
        begin
          if Integer(cbPerfil.Items.Objects[I]) = IdPerfilUser then
          begin
            cbPerfil.ItemIndex := I;
            Break;
          end;
        end;
      end;
    finally
      Qry.Free;
    end;
  end;
end;

procedure TfrmCadastroUsuario.btnCancelarClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfrmCadastroUsuario.btnSalvarClick(Sender: TObject);
var
  Qry: TFDQuery;
  HashSenha, AtivoStr: string;
  IdPerfilSel: Integer;
begin
  if Trim(edtNome.Text) = '' then
  begin
    ShowMessage('Informe o nome.');
    edtNome.SetFocus;
    Exit;
  end;

  if cbPerfil.ItemIndex < 0 then
  begin
    ShowMessage('Selecione um perfil de acesso.');
    cbPerfil.SetFocus;
    Exit;
  end;

  if Trim(edtLogin.Text) = '' then
  begin
    ShowMessage('Informe o login.');
    edtLogin.SetFocus;
    Exit;
  end;

  if (FIdUsuario = 0) and (Trim(edtSenha.Text) = '') then
  begin
    ShowMessage('Informe a senha.');
    edtSenha.SetFocus;
    Exit;
  end;

  IdPerfilSel := Integer(cbPerfil.Items.Objects[cbPerfil.ItemIndex]);

  AtivoStr := 'N';
  if chkAtivo.Checked then
    AtivoStr := 'S';

  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := FConnection;

    if FIdUsuario = 0 then
    begin
      // Inclusao
      HashSenha := THashSHA2.GetHashString(edtSenha.Text);
      Qry.SQL.Text := 'INSERT INTO USUARIO (ID_PERFIL, NOME, LOGIN, SENHA, ATIVO) ' +
                      'VALUES (:ID_PERFIL, :NOME, :LOGIN, :SENHA, :ATIVO)';
      Qry.ParamByName('ID_PERFIL').AsInteger := IdPerfilSel;
      Qry.ParamByName('NOME').AsString       := edtNome.Text;
      Qry.ParamByName('LOGIN').AsString      := edtLogin.Text;
      Qry.ParamByName('SENHA').AsString      := HashSenha;
      Qry.ParamByName('ATIVO').AsString      := AtivoStr;
      Qry.ExecSQL;
    end
    else
    begin
      // Alteracao
      if Trim(edtSenha.Text) <> '' then
      begin
        HashSenha := THashSHA2.GetHashString(edtSenha.Text);
        Qry.SQL.Text := 'UPDATE USUARIO SET ID_PERFIL = :ID_PERFIL, NOME = :NOME, LOGIN = :LOGIN, SENHA = :SENHA, ATIVO = :ATIVO WHERE ID_USUARIO = :ID';
        Qry.ParamByName('SENHA').AsString := HashSenha;
      end
      else
      begin
        Qry.SQL.Text := 'UPDATE USUARIO SET ID_PERFIL = :ID_PERFIL, NOME = :NOME, LOGIN = :LOGIN, ATIVO = :ATIVO WHERE ID_USUARIO = :ID';
      end;

      Qry.ParamByName('ID_PERFIL').AsInteger := IdPerfilSel;
      Qry.ParamByName('NOME').AsString       := edtNome.Text;
      Qry.ParamByName('LOGIN').AsString      := edtLogin.Text;
      Qry.ParamByName('ATIVO').AsString      := AtivoStr;
      Qry.ParamByName('ID').AsInteger         := FIdUsuario;
      Qry.ExecSQL;
    end;

    ShowMessage('Usuário salvo com sucesso!');
    ModalResult := mrOk;
  finally
    Qry.Free;
  end;
end;

end.