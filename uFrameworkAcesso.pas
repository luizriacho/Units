unit uFrameworkAcesso;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Vcl.Forms,
  Vcl.Controls, System.Rtti, FireDAC.Comp.Client, FireDAC.Stan.Param,
  System.Hash;

type
  TPermissaoModulo = record
    CanAccess: Boolean;
    CanInsert: Boolean;
    CanEdit: Boolean;
    CanDelete: Boolean;
  end;

  TPermissaoComponente = record
    Habilitado: Boolean;
    Visivel: Boolean;
  end;

  TSessaoUsuario = class
  private
    FNomePerfil: string;
    FIdUsuario: Integer;
    FIdPerfil: Integer;
    FNome: string;
    FLogin: string;
    FPermissoes: TDictionary<string, TPermissaoModulo>;
    FPermissoesComponentes: TDictionary<string, TPermissaoComponente>;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LimparSessao;
    procedure AdicionarPermissao(const ANomeModulo: string;
      const APermissao: TPermissaoModulo);
    procedure AdicionarPermissaoComponente(const AForm, AComponente: string;
      const APerm: TPermissaoComponente);

    function TemPermissaoAcesso(const ANomeModulo: string): Boolean;
    function ObterPermissao(const ANomeModulo: string): TPermissaoModulo;
    function ObterPermissaoComponente(const AForm, AComponente: string;
      out APerm: TPermissaoComponente): Boolean;

    property IdUsuario: Integer read FIdUsuario write FIdUsuario;
    property IdPerfil: Integer read FIdPerfil write FIdPerfil;
    property Nome: string read FNome write FNome;
    property Login: string read FLogin write FLogin;
    property NomePerfil: string read FNomePerfil write FNomePerfil;
  end;

  TControleAcessoEngine = class
  public
    { Autentica o usuario e carrega as permissoes na sessao }
    class function AutenticarECarregarSessao(AConnection: TFDConnection;
      const AUsuario, ASenha: string): Boolean;

    { Aplica as permissoes em qualquer Form via RTTI }
    class procedure AplicarPermissoesForm(AForm: TForm);
  end;

var
  Sessao: TSessaoUsuario;

implementation

{ TSessaoUsuario }

constructor TSessaoUsuario.Create;
begin
  FPermissoes := TDictionary<string, TPermissaoModulo>.Create;
  FPermissoesComponentes := TDictionary<string, TPermissaoComponente>.Create;
  LimparSessao;
end;

destructor TSessaoUsuario.Destroy;
begin
  FPermissoes.Free;
  FPermissoesComponentes.Free;
  inherited;
end;

procedure TSessaoUsuario.LimparSessao;
begin
  FIdUsuario := 0;
  FIdPerfil := 0;
  FNome := '';
  FLogin := '';
  FNomePerfil := '';
  FPermissoes.Clear;
  FPermissoesComponentes.Clear;
end;

procedure TSessaoUsuario.AdicionarPermissao(const ANomeModulo: string;
  const APermissao: TPermissaoModulo);
begin
  FPermissoes.AddOrSetValue(UpperCase(ANomeModulo), APermissao);
end;

procedure TSessaoUsuario.AdicionarPermissaoComponente(const AForm,
  AComponente: string; const APerm: TPermissaoComponente);
var
  Chave: string;
begin
  Chave := UpperCase(AForm + '.' + AComponente);
  FPermissoesComponentes.AddOrSetValue(Chave, APerm);
end;

function TSessaoUsuario.TemPermissaoAcesso(const ANomeModulo: string): Boolean;
var
  Perm: TPermissaoModulo;
begin
  Result := False;
  if FPermissoes.TryGetValue(UpperCase(ANomeModulo), Perm) then
    Result := Perm.CanAccess;
end;

function TSessaoUsuario.ObterPermissao(const ANomeModulo: string)
  : TPermissaoModulo;
begin
  if not FPermissoes.TryGetValue(UpperCase(ANomeModulo), Result) then
  begin
    Result.CanAccess := False;
    Result.CanInsert := False;
    Result.CanEdit := False;
    Result.CanDelete := False;
  end;
end;

function TSessaoUsuario.ObterPermissaoComponente(const AForm,
  AComponente: string; out APerm: TPermissaoComponente): Boolean;
var
  Chave: string;
begin
  Chave := UpperCase(AForm + '.' + AComponente);
  Result := FPermissoesComponentes.TryGetValue(Chave, APerm);
end;

{ TControleAcessoEngine }

class function TControleAcessoEngine.AutenticarECarregarSessao
  (AConnection: TFDConnection; const AUsuario, ASenha: string): Boolean;
var
  Qry: TFDQuery;
  HashSenha: string;
  Perm: TPermissaoModulo;
  PermComp: TPermissaoComponente;
begin
  Result := False;
  if (AConnection = nil) or not AConnection.Connected then
    Exit;

  HashSenha := THashSHA2.GetHashString(ASenha);

  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := AConnection;

    // 1. Valida Usuario, Senha e traz os dados do Perfil via LEFT JOIN
    Qry.SQL.Text :=
      'SELECT U.ID_USUARIO, U.ID_PERFIL, U.NOME, U.LOGIN, P.NOME_PERFIL ' +
      'FROM USUARIO U ' +
      'LEFT JOIN PERFIL P ON (P.ID_PERFIL = U.ID_PERFIL) ' +
      'WHERE UPPER(U.LOGIN) = UPPER(:LOGIN) AND U.SENHA = :SENHA AND U.ATIVO = ''S''';
    Qry.ParamByName('LOGIN').AsString := AUsuario;
    Qry.ParamByName('SENHA').AsString := HashSenha;
    Qry.Open;

    if Qry.IsEmpty then
      Exit;

    Sessao.LimparSessao;
    Sessao.IdUsuario := Qry.FieldByName('ID_USUARIO').AsInteger;
    Sessao.Nome      := Qry.FieldByName('NOME').AsString;
    Sessao.Login     := Qry.FieldByName('LOGIN').AsString;

    if Qry.FindField('ID_PERFIL') <> nil then
      Sessao.IdPerfil := Qry.FieldByName('ID_PERFIL').AsInteger;

    if (Qry.FindField('NOME_PERFIL') <> nil) and not Qry.FieldByName('NOME_PERFIL').IsNull then
      Sessao.NomePerfil := Qry.FieldByName('NOME_PERFIL').AsString
    else
      Sessao.NomePerfil := 'GERAL';

    // 2. Carrega Permissoes por Modulo do Perfil
    Qry.Close;
    Qry.SQL.Text :=
      'SELECT M.NOME_MODULO, P.CAN_ACCESS, P.CAN_INSERT, P.CAN_EDIT, P.CAN_DELETE ' +
      'FROM PERMISSAO_PERFIL P ' +
      'INNER JOIN MODULO M ON (M.ID_MODULO = P.ID_MODULO) ' +
      'WHERE P.ID_PERFIL = :ID_PERFIL';
    Qry.ParamByName('ID_PERFIL').AsInteger := Sessao.IdPerfil;
    Qry.Open;

    while not Qry.Eof do
    begin
      Perm.CanAccess := Qry.FieldByName('CAN_ACCESS').AsString = 'S';
      Perm.CanInsert := Qry.FieldByName('CAN_INSERT').AsString = 'S';
      Perm.CanEdit   := Qry.FieldByName('CAN_EDIT').AsString = 'S';
      Perm.CanDelete := Qry.FieldByName('CAN_DELETE').AsString = 'S';

      Sessao.AdicionarPermissao(Qry.FieldByName('NOME_MODULO').AsString, Perm);
      Qry.Next;
    end;

    // 3. Carrega Permissoes por Componentes Especificos do Perfil
    Qry.Close;
    Qry.SQL.Text :=
      'SELECT NOME_FORMULARIO, NOME_COMPONENTE, HABILITADO, VISIVEL ' +
      'FROM PERMISSAO_COMPONENTE ' +
      'WHERE ID_PERFIL = :ID_PERFIL';
    Qry.ParamByName('ID_PERFIL').AsInteger := Sessao.IdPerfil;
    Qry.Open;

    while not Qry.Eof do
    begin
      PermComp.Habilitado := Qry.FieldByName('HABILITADO').AsString = 'S';
      PermComp.Visivel    := Qry.FieldByName('VISIVEL').AsString = 'S';

      Sessao.AdicionarPermissaoComponente(Qry.FieldByName('NOME_FORMULARIO').AsString,
        Qry.FieldByName('NOME_COMPONENTE').AsString, PermComp);
      Qry.Next;
    end;

    Result := True;
  finally
    Qry.Free;
  end;
end;

class procedure TControleAcessoEngine.AplicarPermissoesForm(AForm: TForm);
var
  Contexto: TRttiContext;
  TipoRtti: TRttiType;
  Propriedade: TRttiProperty;
  I: Integer;
  Comp: TComponent;
  PermComp: TPermissaoComponente;
begin
  if (AForm = nil) or (UpperCase(Sessao.Login) = 'ADMIN') or (UpperCase(Sessao.NomePerfil) = 'ADMINISTRADOR') then
    Exit;

  Contexto := TRttiContext.Create;
  try
    for I := 0 to AForm.ComponentCount - 1 do
    begin
      Comp := AForm.Components[I];

      if Sessao.ObterPermissaoComponente(AForm.Name, Comp.Name, PermComp) then
      begin
        TipoRtti := Contexto.GetType(Comp.ClassType);

        // Altera Enabled via RTTI se o componente possuir a propriedade
        Propriedade := TipoRtti.GetProperty('Enabled');
        if (Propriedade <> nil) and Propriedade.IsWritable then
          Propriedade.SetValue(Comp, PermComp.Habilitado);

        // Altera Visible via RTTI se o componente possuir a propriedade
        Propriedade := TipoRtti.GetProperty('Visible');
        if (Propriedade <> nil) and Propriedade.IsWritable then
          Propriedade.SetValue(Comp, PermComp.Visivel);
      end;
    end;
  finally
    Contexto.Free;
  end;
end;

initialization

Sessao := TSessaoUsuario.Create;

finalization

Sessao.Free;

end.
