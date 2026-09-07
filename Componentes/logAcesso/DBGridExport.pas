unit DBGridExport;

interface

uses
  Winapi.Windows
, System.SysUtils
, System.Classes
, System.Types
, Vcl.Controls
, Vcl.ExtCtrls
, Vcl.StdCtrls
, Vcl.Buttons
, Vcl.Grids
, Vcl.DBGrids
, Vcl.Graphics
, Vcl.Printers
, Data.DB
, FireDAC.Comp.Client
, FireDAC.Stan.Option
, Vcl.Dialogs
, UExcelService;

type
  TDBGridExport = class(TPanel)
  private
    FDataSource: TDataSource;

    // Controles internos
    FPanelBotoes: TPanel;
    FPanelGrid: TPanel;
    FGrid: TDBGrid;

    FBtnExportarExcel: TBitBtn;
    FBtnImprimir: TBitBtn;

    FExibirExportarExcel: Boolean;
    FExibirImprimir: Boolean;
    FTituloRelatorio: string;
    FTagExportacao: Integer;

    procedure SetDataSource(const Value: TDataSource);
    procedure SetExibirExportarExcel(const Value: Boolean);
    procedure SetExibirImprimir(const Value: Boolean);

    procedure CriarControles;
    procedure AjustarVisibilidadeBotoes;
    procedure OnBtnExcelClick(Sender: TObject);
    procedure OnBtnImprimirClick(Sender: TObject);
    procedure OnGridTitleClick(Column: TColumn);
    procedure OnGridDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure OrdenarGrid(Column: TColumn);
    procedure ImprimirGrid;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure CarregarEAutoAjustarColunas;
    property Grid: TDBGrid read FGrid;
  published
    property DataSource: TDataSource read FDataSource write SetDataSource;
    property ExibirExportarExcel: Boolean read FExibirExportarExcel write SetExibirExportarExcel default True;
    property ExibirImprimir: Boolean read FExibirImprimir write SetExibirImprimir default True;
    property TituloRelatorio: string read FTituloRelatorio write FTituloRelatorio;
    property TagExportacao: Integer read FTagExportacao write FTagExportacao default 0;

    property Align;
    property Anchors;
    property Visible;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Meus Componentes', [TDBGridExport]);
end;

{ TDBGridExport }

constructor TDBGridExport.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Width := 600;
  Height := 400;
  Caption := '';
  BevelOuter := bvNone;

  FExibirExportarExcel := True;
  FExibirImprimir := True;
  FTituloRelatorio := 'Relatório de Dados';
  FTagExportacao := 0;

  CriarControles;
end;

destructor TDBGridExport.Destroy;
begin
  inherited Destroy;
end;

procedure TDBGridExport.SetDataSource(const Value: TDataSource);
begin
  if FDataSource <> Value then
  begin
    FDataSource := Value;
    if Assigned(FGrid) then
    begin
      FGrid.DataSource := FDataSource;
      CarregarEAutoAjustarColunas;
    end;
  end;
end;

procedure TDBGridExport.SetExibirExportarExcel(const Value: Boolean);
begin
  if FExibirExportarExcel <> Value then
  begin
    FExibirExportarExcel := Value;
    AjustarVisibilidadeBotoes;
  end;
end;

procedure TDBGridExport.SetExibirImprimir(const Value: Boolean);
begin
  if FExibirImprimir <> Value then
  begin
    FExibirImprimir := Value;
    AjustarVisibilidadeBotoes;
  end;
end;

procedure TDBGridExport.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FDataSource) then
  begin
    FDataSource := nil;
    if Assigned(FGrid) then
    begin
      FGrid.DataSource := nil;
    end;
  end;
end;

procedure TDBGridExport.CriarControles;
begin
  // Panel Superior para Botões
  FPanelBotoes := TPanel.Create(Self);
  FPanelBotoes.Parent := Self;
  FPanelBotoes.Align := alTop;
  FPanelBotoes.Height := 38;
  FPanelBotoes.BevelOuter := bvLowered;
  FPanelBotoes.Caption := '';

  // Botão Exportar Excel
  FBtnExportarExcel := TBitBtn.Create(Self);
  FBtnExportarExcel.Parent := FPanelBotoes;
  FBtnExportarExcel.Left := 8;
  FBtnExportarExcel.Top := 6;
  FBtnExportarExcel.Width := 110;
  FBtnExportarExcel.Height := 26;
  FBtnExportarExcel.Caption := 'Exportar Excel';
  FBtnExportarExcel.OnClick := OnBtnExcelClick;

  // Botão Imprimir
  FBtnImprimir := TBitBtn.Create(Self);
  FBtnImprimir.Parent := FPanelBotoes;
  FBtnImprimir.Left := 124;
  FBtnImprimir.Top := 6;
  FBtnImprimir.Width := 100;
  FBtnImprimir.Height := 26;
  FBtnImprimir.Caption := 'Imprimir / PDF';
  FBtnImprimir.OnClick := OnBtnImprimirClick;

  // Panel para Grid
  FPanelGrid := TPanel.Create(Self);
  FPanelGrid.Parent := Self;
  FPanelGrid.Align := alClient;
  FPanelGrid.BevelOuter := bvNone;

  // DBGrid
  FGrid := TDBGrid.Create(Self);
  FGrid.Parent := FPanelGrid;
  FGrid.Align := alClient;
  FGrid.Options := [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgTitleClick];
  FGrid.ReadOnly := True;

  FGrid.OnTitleClick := OnGridTitleClick;
  FGrid.OnDrawColumnCell := OnGridDrawColumnCell;

  AjustarVisibilidadeBotoes;
end;

procedure TDBGridExport.AjustarVisibilidadeBotoes;
var
  vPosEsquerda: Integer;
begin
  vPosEsquerda := 8;

  FBtnExportarExcel.Visible := FExibirExportarExcel;
  if FExibirExportarExcel then
  begin
    FBtnExportarExcel.Left := vPosEsquerda;
    vPosEsquerda := vPosEsquerda + FBtnExportarExcel.Width + 6;
  end;

  FBtnImprimir.Visible := FExibirImprimir;
  if FExibirImprimir then
  begin
    FBtnImprimir.Left := vPosEsquerda;
  end;

  FPanelBotoes.Visible := FExibirExportarExcel or FExibirImprimir;
end;

procedure TDBGridExport.CarregarEAutoAjustarColunas;
const
  MARGEM_TITULO = 14;
  MARGEM_DADOS  = 10;
  MAX_LINHAS_AMOSTRA = 500;
var
  i, vLarguraTitulo, vLarguraDados, vLarguraTextoDados, vLarguraFinal: Integer;
  vField: TField;
  vColuna: TColumn;
  vDataSet: TDataSet;
  vBookmark: TBookmark;
  vContadorLinhas: Integer;
  vLarguraTitulos: array of Integer;
  vLarguraColunas: array of Integer;
  vIndiceColuna: Integer;
begin
  if not Assigned(FDataSource) or not Assigned(FDataSource.DataSet) then Exit;

  vDataSet := FDataSource.DataSet;

  // Força limpar colunas anteriores gravadas no DFM/IDE
  FGrid.Columns.Clear;
  FGrid.Canvas.Font := FGrid.Font;

  // 1. Recria as colunas exclusivamente dos TFields que têm Tag = FTagExportacao (0)
  //    e calcula a largura do título (fonte em negrito, igual ao cabeçalho do grid)
  FGrid.Canvas.Font.Style := [fsBold];
  for i := 0 to vDataSet.FieldCount - 1 do
  begin
    vField := vDataSet.Fields[i];
    if (vField.Tag = FTagExportacao) and vField.Visible then
    begin
      vColuna := FGrid.Columns.Add;
      vColuna.FieldName := vField.FieldName;
      vColuna.Title.Caption := vField.DisplayLabel;
      vColuna.Alignment := vField.Alignment;

      SetLength(vLarguraTitulos, FGrid.Columns.Count);

      // Largura do título (com a fonte em negrito usada no cabeçalho do DBGrid)
      vLarguraTitulo := FGrid.Canvas.TextWidth(vField.DisplayLabel) + MARGEM_TITULO;
      vColuna.Width := vLarguraTitulo;
      vLarguraTitulos[FGrid.Columns.Count - 1] := vLarguraTitulo;
    end;
  end;
  FGrid.Canvas.Font.Style := [];

  // 2. Percorre os dados (fonte normal, igual às células) medindo o texto de cada campo
  //    e mantém, por coluna, a maior largura encontrada
  SetLength(vLarguraColunas, FGrid.Columns.Count);
  for vIndiceColuna := 0 to FGrid.Columns.Count - 1 do
    vLarguraColunas[vIndiceColuna] := vLarguraTitulos[vIndiceColuna];

  if not vDataSet.IsEmpty then
  begin
    vDataSet.DisableControls;
    vBookmark := vDataSet.Bookmark;
    try
      vDataSet.First;
      vContadorLinhas := 0;
      while (not vDataSet.Eof) and (vContadorLinhas < MAX_LINHAS_AMOSTRA) do
      begin
        for vIndiceColuna := 0 to FGrid.Columns.Count - 1 do
        begin
          vColuna := FGrid.Columns[vIndiceColuna];
          if Assigned(vColuna.Field) then
          begin
            vLarguraTextoDados := FGrid.Canvas.TextWidth(vColuna.Field.DisplayText) + MARGEM_DADOS;
            if vLarguraTextoDados > vLarguraColunas[vIndiceColuna] then
              vLarguraColunas[vIndiceColuna] := vLarguraTextoDados;
          end;
        end;

        Inc(vContadorLinhas);
        vDataSet.Next;
      end;
    finally
      if vDataSet.BookmarkValid(vBookmark) then
        vDataSet.GotoBookmark(vBookmark);
      vDataSet.FreeBookmark(vBookmark);
      vDataSet.EnableControls;
    end;
  end;

  // 3. Aplica a largura final: título se for maior, senão a largura dos dados
  for vIndiceColuna := 0 to FGrid.Columns.Count - 1 do
  begin
    vLarguraTitulo := vLarguraTitulos[vIndiceColuna];
    vLarguraDados := vLarguraColunas[vIndiceColuna];

    if vLarguraTitulo > vLarguraDados then
      vLarguraFinal := vLarguraTitulo
    else
      vLarguraFinal := vLarguraDados;

    FGrid.Columns[vIndiceColuna].Width := vLarguraFinal;
  end;
end;
procedure TDBGridExport.OnBtnExcelClick(Sender: TObject);
var
  vTituloExportacao: string;
begin
  if not Assigned(FDataSource) or not Assigned(FDataSource.DataSet) then
  begin
    ShowMessage('Não há fonte de dados conectada para exportação.');
    Exit;
  end;

  if FDataSource.DataSet.IsEmpty then
  begin
    ShowMessage('O conjunto de dados está vazio.');
    Exit;
  end;

  CarregarEAutoAjustarColunas;

  if Trim(FTituloRelatorio) <> '' then
    vTituloExportacao := FTituloRelatorio
  else
    vTituloExportacao := 'Relatório de Dados';

  TExcelService.ExportarDataSet(FDataSource.DataSet, FTagExportacao, vTituloExportacao);
end;

procedure TDBGridExport.OnBtnImprimirClick(Sender: TObject);
begin
  ImprimirGrid;
end;

procedure TDBGridExport.OnGridTitleClick(Column: TColumn);
begin
  OrdenarGrid(Column);
end;

procedure TDBGridExport.OrdenarGrid(Column: TColumn);
var
  vDataset: TDataSet;
  vFieldName: string;
  vFDQuery: TFDQuery;
begin
  if not Assigned(Column) or not Assigned(Column.Field) then Exit;

  vDataset := Column.Field.DataSet;
  vFieldName := Column.Field.FieldName;

  if vDataset is TFDQuery then
  begin
    vFDQuery := TFDQuery(vDataset);
    try
      vFDQuery.DisableControls;
      try
        if vFDQuery.FetchOptions.Mode <> fmAll then
        begin
          vFDQuery.FetchOptions.Mode := fmAll;
          vFDQuery.FetchAll;
        end;

        if SameText(vFDQuery.IndexFieldNames, vFieldName) then
        begin
          vFDQuery.IndexFieldNames := vFieldName + ':D';
        end
        else
        begin
          vFDQuery.IndexFieldNames := vFieldName;
        end;

        vFDQuery.First;
      finally
        vFDQuery.EnableControls;
      end;
    except
      ShowMessage('Não foi possível ordenar pelo campo: ' + Column.Title.Caption);
    end;
  end;
end;

procedure TDBGridExport.OnGridDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if not (gdSelected in State) then
  begin
    if Assigned(FDataSource) and Assigned(FDataSource.DataSet) and Odd(FDataSource.DataSet.RecNo) then
      FGrid.Canvas.Brush.Color := clWindow
    else
      FGrid.Canvas.Brush.Color := $00F4F4F4;

    FGrid.Canvas.Font.Color := clWindowText;
  end;

  FGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TDBGridExport.ImprimirGrid;
var
  i, vPosY, vPosX: Integer;
  vAlturaLinha: Integer;
  vScaleX, vScaleY: Double;
  vLarguraColuna: Integer;
  vBookmark: TBookmark;
begin
  if not Assigned(FDataSource) or not Assigned(FDataSource.DataSet) or FDataSource.DataSet.IsEmpty then
  begin
    ShowMessage('Não há dados disponíveis para impressão.');
    Exit;
  end;

  Printer.Orientation := poLandscape;
  Printer.Title := FTituloRelatorio;
  Printer.BeginDoc;
  try
    vScaleX := GetDeviceCaps(Printer.Handle, LOGPIXELSX) / 96.0;
    vScaleY := GetDeviceCaps(Printer.Handle, LOGPIXELSY) / 96.0;

    vAlturaLinha := Round(22 * vScaleY);
    vPosY := Round(40 * vScaleY);

    Printer.Canvas.Font.Name := 'Arial';
    Printer.Canvas.Font.Size := 14;
    Printer.Canvas.Font.Style := [fsBold];
    Printer.Canvas.TextOut(Round(30 * vScaleX), vPosY, FTituloRelatorio);

    vPosY := vPosY + Round(35 * vScaleY);

    Printer.Canvas.Font.Size := 9;
    Printer.Canvas.Font.Style := [fsBold];

    vPosX := Round(30 * vScaleX);
    for i := 0 to FGrid.Columns.Count - 1 do
    begin
      if FGrid.Columns[i].Visible then
      begin
        vLarguraColuna := Round((FGrid.Columns[i].Width + 15) * vScaleX);
        Printer.Canvas.TextOut(vPosX, vPosY, FGrid.Columns[i].Title.Caption);
        vPosX := vPosX + vLarguraColuna;
      end;
    end;

    vPosY := vPosY + vAlturaLinha;

    Printer.Canvas.Pen.Width := Round(1 * vScaleY);
    Printer.Canvas.MoveTo(Round(30 * vScaleX), vPosY);
    Printer.Canvas.LineTo(vPosX, vPosY);
    vPosY := vPosY + Round(8 * vScaleY);

    Printer.Canvas.Font.Style := [];
    FDataSource.DataSet.DisableControls;
    vBookmark := FDataSource.DataSet.Bookmark;
    try
      FDataSource.DataSet.First;
      while not FDataSource.DataSet.Eof do
      begin
        vPosX := Round(30 * vScaleX);
        for i := 0 to FGrid.Columns.Count - 1 do
        begin
          if FGrid.Columns[i].Visible then
          begin
            vLarguraColuna := Round((FGrid.Columns[i].Width + 15) * vScaleX);
            if Assigned(FGrid.Columns[i].Field) then
              Printer.Canvas.TextOut(vPosX, vPosY, FGrid.Columns[i].Field.DisplayText);
            vPosX := vPosX + vLarguraColuna;
          end;
        end;

        vPosY := vPosY + vAlturaLinha;

        if vPosY > (Printer.PageHeight - Round(50 * vScaleY)) then
        begin
          Printer.NewPage;
          vPosY := Round(40 * vScaleY);
        end;

        FDataSource.DataSet.Next;
      end;
    finally
      if FDataSource.DataSet.BookmarkValid(vBookmark) then
        FDataSource.DataSet.GotoBookmark(vBookmark);
      FDataSource.DataSet.FreeBookmark(vBookmark);
      FDataSource.DataSet.EnableControls;
    end;

  finally
    Printer.EndDoc;
  end;
end;

end.
