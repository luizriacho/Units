unit uControleAcessoComponent;

interface

uses
  Winapi.Windows
, System.SysUtils
, System.Classes
, Vcl.Forms
, Vcl.Controls
, Vcl.Dialogs
, FireDAC.Comp.Client
, uFrameworkAcesso
, ufrmLoginGenerico;

type
  [ComponentPlatformsAttribute(pidWin32 or pidWin64)]
  TControleAcessoComponent = class(TComponent)
  private
    FConnection: TFDConnection;
    FAutoLogin: Boolean;
    FUsuarioPadrao: string;
    FSenhaPadrao: string;
    FApenasEmDebug: Boolean;
    FAutoCriarTabelas: Boolean;
    procedure SetConnection(const Value: TFDConnection);
    procedure VerificarECriarEstruturaBanco;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function ExecutarAutenticacao: Boolean;
  published
    property Connection: TFDConnection read FConnection write SetConnection;
    property AutoLogin: Boolean read FAutoLogin write FAutoLogin;
    property UsuarioPadrao: string read FUsuarioPadrao write FUsuarioPadrao;
    property SenhaPadrao: string read FSenhaPadrao write FSenhaPadrao;
    property ApenasEmDebug: Boolean read FApenasEmDebug write FApenasEmDebug;
    property AutoCriarTabelas: Boolean read FAutoCriarTabelas write FAutoCriarTabelas default True;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Meus Componentes', [TControleAcessoComponent]);
end;

{ TControleAcessoComponent }

constructor TControleAcessoComponent.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FAutoLogin := False;
  FUsuarioPadrao := 'luiz';
  FSenhaPadrao := '1234';
  FApenasEmDebug := True;
  FAutoCriarTabelas := True;
end;

destructor TControleAcessoComponent.Destroy;
begin
  inherited Destroy;
end;

procedure TControleAcessoComponent.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FConnection) then
    FConnection := nil;
end;

procedure TControleAcessoComponent.SetConnection(const Value: TFDConnection);
begin
  if FConnection <> Value then
  begin
    FConnection := Value;
    if Assigned(FConnection) then
      FConnection.FreeNotification(Self);
  end;
end;

procedure TControleAcessoComponent.VerificarECriarEstruturaBanco;
var
  vQuery: TFDQuery;
begin
  if not Assigned(FConnection) or not FConnection.Connected then Exit;

  vQuery := TFDQuery.Create(nil);
  try
    vQuery.Connection := FConnection;

    // 1. TABELA DE PERFIS DE ACESSO
    vQuery.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ''PERFIL''';
    vQuery.Open;
    if vQuery.IsEmpty then
    begin
      vQuery.Close;
      vQuery.SQL.Text :=
        'CREATE TABLE PERFIL (' +
        '  ID_PERFIL INTEGER NOT NULL,' +
        '  NOME_PERFIL VARCHAR(50) NOT NULL UNIQUE,' +
        '  CONSTRAINT PK_PERFIL PRIMARY KEY (ID_PERFIL)' +
        ')';
      vQuery.ExecSQL;

      vQuery.SQL.Text := 'CREATE SEQUENCE GEN_PERFIL_ID';
      vQuery.ExecSQL;

      vQuery.SQL.Text :=
        'CREATE TRIGGER BI_PERFIL_ID FOR PERFIL ' +
        'ACTIVE BEFORE INSERT POSITION 0 AS ' +
        'BEGIN ' +
        '  IF (NEW.ID_PERFIL IS NULL OR NEW.ID_PERFIL = 0) THEN ' +
        '    NEW.ID_PERFIL = NEXT VALUE FOR GEN_PERFIL_ID; ' +
        'END';
      vQuery.ExecSQL;

      // Carga inicial dos perfis
      vQuery.SQL.Text := 'INSERT INTO PERFIL (ID_PERFIL, NOME_PERFIL) VALUES (1, ''ADMINISTRADOR'')';
      vQuery.ExecSQL;
      vQuery.SQL.Text := 'INSERT INTO PERFIL (ID_PERFIL, NOME_PERFIL) VALUES (2, ''OPERADOR'')';
      vQuery.ExecSQL;
    end;

    // 2. TABELA DE MÓDULOS
    vQuery.Close;
    vQuery.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ''MODULO''';
    vQuery.Open;
    if vQuery.IsEmpty then
    begin
      vQuery.Close;
      vQuery.SQL.Text :=
        'CREATE TABLE MODULO (' +
        '  ID_MODULO INTEGER NOT NULL,' +
        '  NOME_MODULO VARCHAR(50) NOT NULL UNIQUE,' +
        '  DESCRICAO VARCHAR(100),' +
        '  CONSTRAINT PK_MODULO PRIMARY KEY (ID_MODULO)' +
        ')';
      vQuery.ExecSQL;

      vQuery.SQL.Text := 'CREATE SEQUENCE GEN_MODULO_ID';
      vQuery.ExecSQL;

      vQuery.SQL.Text :=
        'CREATE TRIGGER BI_MODULO_ID FOR MODULO ' +
        'ACTIVE BEFORE INSERT POSITION 0 AS ' +
        'BEGIN ' +
        '  IF (NEW.ID_MODULO IS NULL OR NEW.ID_MODULO = 0) THEN ' +
        '    NEW.ID_MODULO = NEXT VALUE FOR GEN_MODULO_ID; ' +
        'END';
      vQuery.ExecSQL;
    end;

    // 3. TABELA DE USUÁRIOS
    vQuery.Close;
    vQuery.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ''USUARIO''';
    vQuery.Open;
    if vQuery.IsEmpty then
    begin
      vQuery.Close;
      vQuery.SQL.Text :=
        'CREATE TABLE USUARIO (' +
        '  ID_USUARIO INTEGER NOT NULL,' +
        '  ID_PERFIL INTEGER NOT NULL,' +
        '  NOME VARCHAR(100) NOT NULL,' +
        '  LOGIN VARCHAR(30) NOT NULL UNIQUE,' +
        '  SENHA VARCHAR(64) NOT NULL,' +
        '  ATIVO CHAR(1) DEFAULT ''S'' CHECK (ATIVO IN (''S'', ''N'')),' +
        '  CONSTRAINT PK_USUARIO PRIMARY KEY (ID_USUARIO),' +
        '  CONSTRAINT FK_USUARIO_PERFIL FOREIGN KEY (ID_PERFIL) REFERENCES PERFIL(ID_PERFIL)' +
        ')';
      vQuery.ExecSQL;

      vQuery.SQL.Text := 'CREATE SEQUENCE GEN_USUARIO_ID';
      vQuery.ExecSQL;

      vQuery.SQL.Text :=
        'CREATE TRIGGER BI_USUARIO_ID FOR USUARIO ' +
        'ACTIVE BEFORE INSERT POSITION 0 AS ' +
        'BEGIN ' +
        '  IF (NEW.ID_USUARIO IS NULL OR NEW.ID_USUARIO = 0) THEN ' +
        '    NEW.ID_USUARIO = NEXT VALUE FOR GEN_USUARIO_ID; ' +
        'END';
      vQuery.ExecSQL;

      // Cadastra o usuário padrão ADMIN (Senha '123456' em SHA-256)
      vQuery.SQL.Text :=
        'INSERT INTO USUARIO (ID_PERFIL, NOME, LOGIN, SENHA, ATIVO) ' +
        'VALUES (1, ''ADMINISTRADOR'', ''ADMIN'', ''8d969eef6ecad3c29a3a629280e686cf0c3f5d5a86aff3ca12020c923adc6c92'', ''S'')';
      vQuery.ExecSQL;

      // Cadastra o usuário 'luiz' (Senha '1234' em SHA-256) se não existir
      vQuery.SQL.Text :=
        'INSERT INTO USUARIO (ID_PERFIL, NOME, LOGIN, SENHA, ATIVO) ' +
        'VALUES (1, ''LUIZ'', ''luiz'', ''03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4'', ''S'')';
      vQuery.ExecSQL;
    end;

    // 4. TABELA DE PERMISSÕES DO PERFIL POR MÓDULO
    vQuery.Close;
    vQuery.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ''PERMISSAO_PERFIL''';
    vQuery.Open;
    if vQuery.IsEmpty then
    begin
      vQuery.Close;
      vQuery.SQL.Text :=
        'CREATE TABLE PERMISSAO_PERFIL (' +
        '  ID_PERFIL INTEGER NOT NULL,' +
        '  ID_MODULO INTEGER NOT NULL,' +
        '  CAN_ACCESS CHAR(1) DEFAULT ''N'' CHECK (CAN_ACCESS IN (''S'', ''N'')),' +
        '  CAN_INSERT CHAR(1) DEFAULT ''N'' CHECK (CAN_INSERT IN (''S'', ''N'')),' +
        '  CAN_EDIT CHAR(1) DEFAULT ''N'' CHECK (CAN_EDIT IN (''S'', ''N'')),' +
        '  CAN_DELETE CHAR(1) DEFAULT ''N'' CHECK (CAN_DELETE IN (''S'', ''N'')),' +
        '  CONSTRAINT PK_PERMISSAO_PERFIL PRIMARY KEY (ID_PERFIL, ID_MODULO),' +
        '  CONSTRAINT FK_PERM_PERFIL FOREIGN KEY (ID_PERFIL) REFERENCES PERFIL(ID_PERFIL) ON DELETE CASCADE,' +
        '  CONSTRAINT FK_PERM_MOD_PERFIL FOREIGN KEY (ID_MODULO) REFERENCES MODULO(ID_MODULO) ON DELETE CASCADE' +
        ')';
      vQuery.ExecSQL;
    end;

    // 5. TABELA DE PERMISSÕES POR COMPONENTE
    vQuery.Close;
    vQuery.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ''PERMISSAO_COMPONENTE''';
    vQuery.Open;
    if vQuery.IsEmpty then
    begin
      vQuery.Close;
      vQuery.SQL.Text :=
        'CREATE TABLE PERMISSAO_COMPONENTE (' +
        '  ID_PERFIL INTEGER NOT NULL,' +
        '  NOME_FORMULARIO VARCHAR(60) NOT NULL,' +
        '  NOME_COMPONENTE VARCHAR(60) NOT NULL,' +
        '  HABILITADO CHAR(1) DEFAULT ''S'' CHECK (HABILITADO IN (''S'', ''N'')),' +
        '  VISIVEL CHAR(1) DEFAULT ''S'' CHECK (VISIVEL IN (''S'', ''N'')),' +
        '  CONSTRAINT PK_PERMISSAO_COMPONENTE PRIMARY KEY (ID_PERFIL, NOME_FORMULARIO, NOME_COMPONENTE),' +
        '  CONSTRAINT FK_PERM_COMP_PERFIL FOREIGN KEY (ID_PERFIL) REFERENCES PERFIL(ID_PERFIL) ON DELETE CASCADE' +
        ')';
      vQuery.ExecSQL;
    end;

  finally
    vQuery.Free;
  end;
end;

function TControleAcessoComponent.ExecutarAutenticacao: Boolean;
var
  vPodeFazerAutoLogin: Boolean;
begin
  Result := False;

  if not Assigned(FConnection) then
  begin
    ShowMessage('TControleAcessoComponent: Nenhuma conexão TFDConnection atribuída.');
    Exit;
  end;

  if not FConnection.Connected then
    FConnection.Connected := True;

  if FAutoCriarTabelas then
    VerificarECriarEstruturaBanco;

  vPodeFazerAutoLogin := FAutoLogin;

  if vPodeFazerAutoLogin and FApenasEmDebug then
  begin
    {$IFNDEF DEBUG}
      vPodeFazerAutoLogin := False;
    {$ENDIF}
  end;

  if vPodeFazerAutoLogin then
  begin
    Result := TControleAcessoEngine.AutenticarECarregarSessao(FConnection, FUsuarioPadrao, FSenhaPadrao);
    if Result then
      Exit;
  end;

  Result := TfrmLoginGenerico.ExecutarLogin(FConnection);
end;

end.
