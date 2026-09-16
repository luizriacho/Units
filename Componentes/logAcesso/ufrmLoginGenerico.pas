unit ufrmLoginGenerico;

interface

uses
  Winapi.Windows
, Winapi.Messages
, System.SysUtils
, System.Variants
, System.Classes
, Vcl.Graphics
, Vcl.Controls
, Vcl.Forms
, Vcl.Dialogs
, Vcl.StdCtrls
, Vcl.Buttons
, FireDAC.Comp.Client
, uFrameworkAcesso;

// Descomente ou comente a linha abaixo para ativar/desativar o Auto-Login de Desenvolvimento manualmente:
{$DEFINE AUTO_LOGIN_DEV}

type
  TfrmLoginGenerico = class(TForm)
    edtUsuario: TEdit;
    edtSenha: TEdit;
    btnEntrar: TBitBtn;
    btnCancelar: TBitBtn;
    lblUsuario: TLabel;
    lblSenha: TLabel;
    procedure btnEntrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    FConnection: TFDConnection;
  public
    class function ExecutarLogin(AConnection: TFDConnection; const AAutoLoginDev: Boolean = False): Boolean;
  end;

var
  frmLoginGenerico: TfrmLoginGenerico;

const
  // Dados do usuário para Auto Login no ambiente de desenvolvimento
  USUARIO_DEV_AUTO_LOGIN = 'ADMIN';
  SENHA_DEV_AUTO_LOGIN   = '123';

implementation

{$R *.dfm}

class function TfrmLoginGenerico.ExecutarLogin(AConnection: TFDConnection; const AAutoLoginDev: Boolean = False): Boolean;
var
  Frm: TfrmLoginGenerico;
  vAutoLoginAtivo: Boolean;
begin
  vAutoLoginAtivo := AAutoLoginDev;

  // 1. Verifica se a diretiva de compilação ou modo DEBUG do Delphi está ativa
  {$IFDEF AUTO_LOGIN_DEV}
    vAutoLoginAtivo := True;
  {$ENDIF}

  {$IFDEF DEBUG}
    vAutoLoginAtivo := True;
  {$ENDIF}

  // 2. Se o Auto-Login estiver ativado, autentica silenciosamente sem exibir a tela de login
  if vAutoLoginAtivo then
  begin
    Result := TControleAcessoEngine.AutenticarECarregarSessao(AConnection, USUARIO_DEV_AUTO_LOGIN, SENHA_DEV_AUTO_LOGIN);
    if Result then
      Exit;
  end;

  // 3. Caso contrário (ou se o auto-login falhar), exibe a tela de login normalmente
  Frm := TfrmLoginGenerico.Create(nil);
  try
    Frm.FConnection := AConnection;
    Result := Frm.ShowModal = mrOk;
  finally
    Frm.Free;
  end;
end;

procedure TfrmLoginGenerico.btnCancelarClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfrmLoginGenerico.btnEntrarClick(Sender: TObject);
begin
  if Trim(edtUsuario.Text) = '' then
  begin
    ShowMessage('Informe o usuário.');
    edtUsuario.SetFocus;
    Exit;
  end;

  if Trim(edtSenha.Text) = '' then
  begin
    ShowMessage('Informe a senha.');
    edtSenha.SetFocus;
    Exit;
  end;

  // Tenta autenticar
  if TControleAcessoEngine.AutenticarECarregarSessao(FConnection, edtUsuario.Text, edtSenha.Text) then
  begin
    ModalResult := mrOk; // Define sucesso apenas aqui
  end
  else
  begin
    ShowMessage('Usuário ou senha inválidos!');
    edtSenha.Clear;
    edtSenha.SetFocus;
    // NÃO altera o ModalResult aqui para o formulário NÃO fechar com sucesso
  end;
end;

procedure TfrmLoginGenerico.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
    SelectNext(ActiveControl, True, True);
end;

end.
