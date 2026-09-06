unit ufrmLoginGenerico;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, 
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  FireDAC.Comp.Client, uFrameworkAcesso;

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
    class function ExecutarLogin(AConnection: TFDConnection): Boolean;
  end;

var
  frmLoginGenerico: TfrmLoginGenerico;

implementation

{$R *.dfm}

class function TfrmLoginGenerico.ExecutarLogin(AConnection: TFDConnection): Boolean;
var
  Frm: TfrmLoginGenerico;
begin
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