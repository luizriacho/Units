unit LogAcessoGrid;

interface

uses
  Winapi.Windows
, System.SysUtils
, System.Classes
, System.Types
, Vcl.Controls
, Vcl.ExtCtrls
, Vcl.StdCtrls
, Vcl.ComCtrls
, Vcl.Grids
, Vcl.DBGrids
, Vcl.Graphics
, Data.DB
, FireDAC.Comp.Client
, FireDAC.Stan.Param
, Vcl.Dialogs;

type
  TLogAcessoGrid = class;

  // Classe para agrupar as propriedades das Colunas no Object Inspector
  TLogAcessoColunas = class(TPersistent)
  private
    FOwner: TLogAcessoGrid;
    FDataHora: Boolean;
    FUsuario: Boolean;
    FTipo: Boolean;
    FFormulario: Boolean;
    FAcao: Boolean;
    procedure SetDataHora(const Value: Boolean);
    procedure SetUsuario(const Value: Boolean);
    procedure SetTipo(const Value: Boolean);
    procedure SetFormulario(const Value: Boolean);
    procedure SetAcao(const Value: Boolean);
  public
    constructor Create(AOwner: TLogAcessoGrid);
  published
    property DataHora: Boolean read FDataHora write SetDataHora default True;
    property Usuario: Boolean read FUsuario write SetUsuario default True;
    property Tipo: Boolean read FTipo write SetTipo default True;
    property Formulario: Boolean read FFormulario write SetFormulario default True;
    property Acao: Boolean read FAcao write SetAcao default True;
  end;

  // Classe para agrupar a visibilidade dos Filtros no Object Inspector
  TLogAcessoFiltros = class(TPersistent)
  private
    FOwner: TLogAcessoGrid;
    FPeriodo: Boolean;
    FUsuario: Boolean;
    FTipo: Boolean;
    procedure SetPeriodo(const Value: Boolean);
    procedure SetUsuario(const Value: Boolean);
    procedure SetTipo(const Value: Boolean);
  public
    constructor Create(AOwner: TLogAcessoGrid);
  published
    property Periodo: Boolean read FPeriodo write SetPeriodo default True;
    property Usuario: Boolean read FUsuario write SetUsuario default True;
    property Tipo: Boolean read FTipo write SetTipo default True;
  end;

  // Componente Principal
  TLogAcessoGrid = class(TPanel)
  private
    FConnection: TFDConnection;
    FQuery: TFDQuery;
    FDataSource: TDataSource;

    FColunas: TLogAcessoColunas;
    FFiltros: TLogAcessoFiltros;

    // Controles de Interface
    FPanelFiltros: TPanel;
    FPanelGrid: TPanel;
    FGrid: TDBGrid;

    FdtpInicio: TDateTimePicker;
    FdtpFim: TDateTimePicker;
    FedtUsuario: TEdit;
    FedtTipo: TEdit;
    FbtnConsultar: TButton;

    FlblInicio: TLabel;
    FlblFim: TLabel;
    FlblUsuario: TLabel;
    FlblTipo: TLabel;

    procedure SetConnection(const Value: TFDConnection);

    procedure CriarControles;
    procedure AjustarLayoutFiltros;
    procedure FormatGrid;
    procedure OnBtnConsultarClick(Sender: TObject);
    procedure OnFiltroChange(Sender: TObject);
    procedure OnGridTitleClick(Column: TColumn);
    procedure OnGridDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure OrdenarGrid(Column: TColumn);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Consultar;
  published
    property Connection: TFDConnection read FConnection write SetConnection;

    // Propriedades Agrupadas no Object Inspector
    property Colunas: TLogAcessoColunas read FColunas write FColunas;
    property Filtros: TLogAcessoFiltros read FFiltros write FFiltros;

    property Align;
    property Anchors;
    property Visible;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Meus Componentes', [TLogAcessoGrid]);
end;

{ TLogAcessoColunas }

constructor TLogAcessoColunas.Create(AOwner: TLogAcessoGrid);
begin
  inherited Create;
  FOwner := AOwner;
  FDataHora := True;
  FUsuario := True;
  FTipo := True;
  FFormulario := True;
  FAcao := True;
end;

procedure TLogAcessoColunas.SetAcao(const Value: Boolean);
begin
  if FAcao <> Value then
  begin
    FAcao := Value;
    FOwner.FormatGrid;
  end;
end;

procedure TLogAcessoColunas.SetDataHora(const Value: Boolean);
begin
  if FDataHora <> Value then
  begin
    FDataHora := Value;
    FOwner.FormatGrid;
  end;
end;

procedure TLogAcessoColunas.SetFormulario(const Value: Boolean);
begin
  if FFormulario <> Value then
  begin
    FFormulario := Value;
    FOwner.FormatGrid;
  end;
end;

procedure TLogAcessoColunas.SetTipo(const Value: Boolean);
begin
  if FTipo <> Value then
  begin
    FTipo := Value;
    FOwner.FormatGrid;
  end;
end;

procedure TLogAcessoColunas.SetUsuario(const Value: Boolean);
begin
  if FUsuario <> Value then
  begin
    FUsuario := Value;
    FOwner.FormatGrid;
  end;
end;

{ TLogAcessoFiltros }

constructor TLogAcessoFiltros.Create(AOwner: TLogAcessoGrid);
begin
  inherited Create;
  FOwner := AOwner;
  FPeriodo := True;
  FUsuario := True;
  FTipo := True;
end;

procedure TLogAcessoFiltros.SetPeriodo(const Value: Boolean);
begin
  if FPeriodo <> Value then
  begin
    FPeriodo := Value;
    FOwner.AjustarLayoutFiltros;
  end;
end;

procedure TLogAcessoFiltros.SetTipo(const Value: Boolean);
begin
  if FTipo <> Value then
  begin
    FTipo := Value;
    FOwner.AjustarLayoutFiltros;
  end;
end;

procedure TLogAcessoFiltros.SetUsuario(const Value: Boolean);
begin
  if FUsuario <> Value then
  begin
    FUsuario := Value;
    FOwner.AjustarLayoutFiltros;
  end;
end;

{ TLogAcessoGrid }

constructor TLogAcessoGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Width := 700;
  Height := 450;
  Caption := '';
  BevelOuter := bvNone;

  // Instância do agrupamento de propriedades
  FColunas := TLogAcessoColunas.Create(Self);
  FFiltros := TLogAcessoFiltros.Create(Self);

  // Instância dos objetos de acesso a dados
  FQuery := TFDQuery.Create(Self);
  FDataSource := TDataSource.Create(Self);
  FDataSource.DataSet := FQuery;

  // Constrói a interface visual do componente
  CriarControles;
end;

destructor TLogAcessoGrid.Destroy;
begin
  FreeAndNil(FColunas);
  FreeAndNil(FFiltros);
  inherited Destroy;
end;

procedure TLogAcessoGrid.SetConnection(const Value: TFDConnection);
begin
  if FConnection <> Value then
  begin
    FConnection := Value;
    if Assigned(FConnection) then
    begin
      FQuery.Connection := FConnection;
    end;
  end;
end;

procedure TLogAcessoGrid.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FConnection) then
    FConnection := nil;
end;

procedure TLogAcessoGrid.CriarControles;
begin
  // Panel Superior para Filtros
  FPanelFiltros := TPanel.Create(Self);
  FPanelFiltros.Parent := Self;
  FPanelFiltros.Align := alTop;
  FPanelFiltros.Height := 65;
  FPanelFiltros.BevelOuter := bvLowered;
  FPanelFiltros.Caption := '';

  // Data Inicial
  FlblInicio := TLabel.Create(Self);
  FlblInicio.Parent := FPanelFiltros;
  FlblInicio.Caption := 'Data Inicial:';

  FdtpInicio := TDateTimePicker.Create(Self);
  FdtpInicio.Parent := FPanelFiltros;
  FdtpInicio.Width := 95;
  FdtpInicio.Format := 'dd/MM/yyyy';
  FdtpInicio.Date := Date;

  // Data Final
  FlblFim := TLabel.Create(Self);
  FlblFim.Parent := FPanelFiltros;
  FlblFim.Caption := 'Data Final:';

  FdtpFim := TDateTimePicker.Create(Self);
  FdtpFim.Parent := FPanelFiltros;
  FdtpFim.Width := 95;
  FdtpFim.Format := 'dd/MM/yyyy';
  FdtpFim.Date := Date;

  // Usuário
  FlblUsuario := TLabel.Create(Self);
  FlblUsuario.Parent := FPanelFiltros;
  FlblUsuario.Caption := 'Usuário:';

  FedtUsuario := TEdit.Create(Self);
  FedtUsuario.Parent := FPanelFiltros;
  FedtUsuario.Width := 110;
  FedtUsuario.OnChange := OnFiltroChange;

  // Tipo
  FlblTipo := TLabel.Create(Self);
  FlblTipo.Parent := FPanelFiltros;
  FlblTipo.Caption := 'Tipo:';

  FedtTipo := TEdit.Create(Self);
  FedtTipo.Parent := FPanelFiltros;
  FedtTipo.Width := 110;
  FedtTipo.OnChange := OnFiltroChange;

  // Botão Consultar
  FbtnConsultar := TButton.Create(Self);
  FbtnConsultar.Parent := FPanelFiltros;
  FbtnConsultar.Width := 90;
  FbtnConsultar.Height := 25;
  FbtnConsultar.Caption := 'Consultar';
  FbtnConsultar.OnClick := OnBtnConsultarClick;

  // Align Grid
  FPanelGrid := TPanel.Create(Self);
  FPanelGrid.Parent := Self;
  FPanelGrid.Align := alClient;
  FPanelGrid.BevelOuter := bvNone;

  FGrid := TDBGrid.Create(Self);
  FGrid.Parent := FPanelGrid;
  FGrid.Align := alClient;
  FGrid.DataSource := FDataSource;
  FGrid.Options := [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection];
  FGrid.ReadOnly := True;

  FGrid.OnTitleClick := OnGridTitleClick;
  FGrid.OnDrawColumnCell := OnGridDrawColumnCell;

  // Calcula a posição dos campos conforme visibilidade
  AjustarLayoutFiltros;
end;

procedure TLogAcessoGrid.AjustarLayoutFiltros;
var
  vPosEsquerda: Integer;
begin
  vPosEsquerda := 10;

  // Filtro de Período
  FlblInicio.Visible := FFiltros.Periodo;
  FdtpInicio.Visible := FFiltros.Periodo;
  FlblFim.Visible := FFiltros.Periodo;
  FdtpFim.Visible := FFiltros.Periodo;

  if FFiltros.Periodo then
  begin
    FlblInicio.Left := vPosEsquerda;
    FlblInicio.Top := 10;
    FdtpInicio.Left := vPosEsquerda;
    FdtpInicio.Top := 28;

    vPosEsquerda := vPosEsquerda + FdtpInicio.Width + 10;

    FlblFim.Left := vPosEsquerda;
    FlblFim.Top := 10;
    FdtpFim.Left := vPosEsquerda;
    FdtpFim.Top := 28;

    vPosEsquerda := vPosEsquerda + FdtpFim.Width + 15;
  end;

  // Filtro de Usuário
  FlblUsuario.Visible := FFiltros.Usuario;
  FedtUsuario.Visible := FFiltros.Usuario;

  if FFiltros.Usuario then
  begin
    FlblUsuario.Left := vPosEsquerda;
    FlblUsuario.Top := 10;
    FedtUsuario.Left := vPosEsquerda;
    FedtUsuario.Top := 28;

    vPosEsquerda := vPosEsquerda + FedtUsuario.Width + 15;
  end;

  // Filtro de Tipo
  FlblTipo.Visible := FFiltros.Tipo;
  FedtTipo.Visible := FFiltros.Tipo;

  if FFiltros.Tipo then
  begin
    FlblTipo.Left := vPosEsquerda;
    FlblTipo.Top := 10;
    FedtTipo.Left := vPosEsquerda;
    FedtTipo.Top := 28;

    vPosEsquerda := vPosEsquerda + FedtTipo.Width + 15;
  end;

  // Botão Consultar ajusta automaticamente
  FbtnConsultar.Left := vPosEsquerda;
  FbtnConsultar.Top := 26;
end;

procedure TLogAcessoGrid.OnBtnConsultarClick(Sender: TObject);
begin
  Consultar;
end;

procedure TLogAcessoGrid.OnFiltroChange(Sender: TObject);
begin
  Consultar;
end;

procedure TLogAcessoGrid.OnGridTitleClick(Column: TColumn);
begin
  OrdenarGrid(Column);
end;

procedure TLogAcessoGrid.OrdenarGrid(Column: TColumn);
var
  vDataset: TDataSet;
  vFieldName: string;
begin
  if not Assigned(Column.Field) then Exit;

  vDataset := Column.Field.DataSet;
  vFieldName := Column.FieldName;

  if vDataset is TFDQuery then
  begin
    try
      vDataset.DisableControls;
      try
        if TFDQuery(vDataset).IndexFieldNames = vFieldName then
        begin
          TFDQuery(vDataset).IndexFieldNames := vFieldName + ':D';
        end
        else
        begin
          TFDQuery(vDataset).IndexFieldNames := vFieldName;
        end;
        vDataset.First;
      finally
        vDataset.EnableControls;
      end;
    except
      ShowMessage('Não foi possível ordenar por este campo.');
    end;
  end;
end;

procedure TLogAcessoGrid.OnGridDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if not (gdSelected in State) then
  begin
    if Odd(FDataSource.DataSet.RecNo) then
      FGrid.Canvas.Brush.Color := clWindow
    else
      FGrid.Canvas.Brush.Color := $00F4F4F4;

    FGrid.Canvas.Font.Color := clWindowText;
  end;

  FGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TLogAcessoGrid.Consultar;
begin
  if not Assigned(FConnection) then Exit;

  FQuery.Close;
  FQuery.SQL.Clear;
  FQuery.SQL.Add('SELECT DATA_HORA, USUARIO, TIPO, FORMULARIO, ACAO ');
  FQuery.SQL.Add('FROM LOG_ACESSO ');
  FQuery.SQL.Add('WHERE 1=1 ');

  if FFiltros.Periodo then
  begin
    FQuery.SQL.Add('  AND DATA_HORA >= :pDATA_INI ');
    FQuery.SQL.Add('  AND DATA_HORA <= :pDATA_FIM ');
  end;

  if FFiltros.Usuario and (Trim(FedtUsuario.Text) <> '') then
    FQuery.SQL.Add('  AND UPPER(USUARIO) LIKE :pUSUARIO ');

  if FFiltros.Tipo and (Trim(FedtTipo.Text) <> '') then
    FQuery.SQL.Add('  AND UPPER(TIPO) LIKE :pTIPO ');

  FQuery.SQL.Add('ORDER BY DATA_HORA DESC');

  // Atribuição de Parâmetros
  if FFiltros.Periodo then
  begin
    FQuery.ParamByName('pDATA_INI').AsDateTime := Trunc(FdtpInicio.Date);
    FQuery.ParamByName('pDATA_FIM').AsDateTime := Trunc(FdtpFim.Date) + 0.99999;
  end;

  if FFiltros.Usuario and (Trim(FedtUsuario.Text) <> '') then
    FQuery.ParamByName('pUSUARIO').AsString := '%' + UpperCase(Trim(FedtUsuario.Text)) + '%';

  if FFiltros.Tipo and (Trim(FedtTipo.Text) <> '') then
    FQuery.ParamByName('pTIPO').AsString := '%' + UpperCase(Trim(FedtTipo.Text)) + '%';

  FQuery.Open;
  FormatGrid;
end;

procedure TLogAcessoGrid.FormatGrid;
var
  Coluna: TColumn;
begin
  if FQuery.IsEmpty then Exit;

  FGrid.Columns.Clear;

  // Coluna DATA_HORA
  if FColunas.DataHora then
  begin
    Coluna := FGrid.Columns.Add;
    Coluna.FieldName := 'DATA_HORA';
    Coluna.Title.Caption := 'Data/Hora';
    Coluna.Width := 130;
    if Assigned(FQuery.FindField('DATA_HORA')) and (FQuery.FieldByName('DATA_HORA') is TDateTimeField) then
      TDateTimeField(FQuery.FieldByName('DATA_HORA')).DisplayFormat := 'dd/mm/yyyy hh:nn:ss';
  end;

  // Coluna USUARIO
  if FColunas.Usuario then
  begin
    Coluna := FGrid.Columns.Add;
    Coluna.FieldName := 'USUARIO';
    Coluna.Title.Caption := 'Usuário';
    Coluna.Width := 110;
  end;

  // Coluna TIPO
  if FColunas.Tipo then
  begin
    Coluna := FGrid.Columns.Add;
    Coluna.FieldName := 'TIPO';
    Coluna.Title.Caption := 'Tipo';
    Coluna.Width := 90;
  end;

  // Coluna FORMULARIO
  if FColunas.Formulario then
  begin
    Coluna := FGrid.Columns.Add;
    Coluna.FieldName := 'FORMULARIO';
    Coluna.Title.Caption := 'Formulário';
    Coluna.Width := 140;
  end;

  // Coluna ACAO
  if FColunas.Acao then
  begin
    Coluna := FGrid.Columns.Add;
    Coluna.FieldName := 'ACAO';
    Coluna.Title.Caption := 'Ação Realizada';
    Coluna.Width := 200;
  end;
end;

end.
