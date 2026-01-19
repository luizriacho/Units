unit UExcelService;

interface

uses
  System.SysUtils, System.Variants, Data.DB, Winapi.ActiveX, System.Win.ComObj, Vcl.Dialogs;

type
  TExcelService = class
  private
    class procedure FormatarPlanilhaExcel(const Excel: Variant; const LinhaFinal, ColunaFinal: Integer);
  public
    class procedure ExportarDataSet(DataSet: TDataSet; ValorTag: Integer; const TituloCabecalho: string);
  end;

implementation

{ TExcelService }

class procedure TExcelService.ExportarDataSet(DataSet: TDataSet; ValorTag: Integer; const TituloCabecalho: string);
var
  Linha, coluna, ColExcel: Integer;
  planilha, Sheet, Dados, vRange: Variant;
  UltimaColunaPreenchida: Integer;
  TotalRegistros: Integer;
  PathImages: string;
  LogoCliente, LogoSic: string;
  PosicaoEsquerdaSic: Double;
const
  xlCenter = -4108;
begin
  if not Assigned(DataSet) or (DataSet.IsEmpty) then
    Exit;

  PathImages  := ExtractFilePath(ParamStr(0)) + 'images\';
  LogoCliente := PathImages + 'logoCliente.png';
  LogoSic     := PathImages + 'logoSic.png';

  TotalRegistros := DataSet.RecordCount;
  DataSet.DisableControls;
  try
    planilha := CreateOleObject('Excel.Application');
    planilha.WorkBooks.Add;
    planilha.Visible := True;
    planilha.ScreenUpdating := False;
    Sheet := planilha.ActiveWorkbook.ActiveSheet;

    // 1. Conta colunas pela TAG
    UltimaColunaPreenchida := 0;
    for coluna := 0 to DataSet.FieldCount - 1 do
      if DataSet.Fields[coluna].Tag = ValorTag then
        Inc(UltimaColunaPreenchida);

    // 2. Ajuste de altura das linhas 1 e 2 para as logos (Aumento da moldura)
    Sheet.Rows.Item[1].RowHeight := 25;
    Sheet.Rows.Item[2].RowHeight := 25; // Somadas dão 50 de altura para o cabeçalho

    // 3. Preparar Array de Dados
    Dados := VarArrayCreate([1, TotalRegistros + 1, 1, UltimaColunaPreenchida], varVariant);

    // Header
    ColExcel := 1;
    for coluna := 0 to DataSet.FieldCount - 1 do
      if DataSet.Fields[coluna].Tag = ValorTag then
      begin
        Dados[1, ColExcel] := DataSet.Fields[coluna].DisplayLabel;
        Inc(ColExcel);
      end;

    // Dados
    Linha := 2;
    DataSet.First;
    while not DataSet.Eof do
    begin
      ColExcel := 1;
      for coluna := 0 to DataSet.FieldCount - 1 do
        if DataSet.Fields[coluna].Tag = ValorTag then
        begin
          if DataSet.Fields[coluna].IsNull then
            Dados[Linha, ColExcel] := ''
          else if DataSet.Fields[coluna] is TNumericField then
            Dados[Linha, ColExcel] := DataSet.Fields[coluna].AsFloat
          else
            Dados[Linha, ColExcel] := DataSet.Fields[coluna].Value;
          Inc(ColExcel);
        end;
      Inc(Linha);
      DataSet.Next;
    end;

    // 4. Despeja os dados
    Sheet.Range[Sheet.Cells[3, 1], Sheet.Cells[3 + TotalRegistros, UltimaColunaPreenchida]].Value := Dados;

    // 5. Formata a Planilha (AutoFit)
    FormatarPlanilhaExcel(planilha, 3 + TotalRegistros, UltimaColunaPreenchida);

    // 6. Configura o Título (Centralizado H e V)
    vRange := Sheet.Range[Sheet.Cells[1, 1], Sheet.Cells[2, UltimaColunaPreenchida]];
    vRange.Merge;
    vRange.Value := TituloCabecalho;
    vRange.Font.Bold := True;
    vRange.Font.Size := 14;
    vRange.HorizontalAlignment := xlCenter;
    vRange.VerticalAlignment := xlCenter;

    // 7. INSERÇÃO DAS LOGOMARCAS (Com ajuste de centralização na altura)
    // Logo Cliente (Esquerda)
    if FileExists(LogoCliente) then
      Sheet.Shapes.AddPicture(LogoCliente, False, True, 5, 5, 65, 38);

    // Logo SIC (Direita) - Posição recalculada após AutoFit
    if FileExists(LogoSic) then
    begin
      PosicaoEsquerdaSic := Sheet.Cells[1, UltimaColunaPreenchida].Left +
                            Sheet.Cells[1, UltimaColunaPreenchida].Width - 65;

      Sheet.Shapes.AddPicture(LogoSic, False, True, PosicaoEsquerdaSic, 5, 42, 42);
    end;

  finally
    DataSet.EnableControls;
    planilha.ScreenUpdating := True;
  end;
end;

class procedure TExcelService.FormatarPlanilhaExcel(const Excel: Variant; const LinhaFinal, ColunaFinal: Integer);
var
  planilha, Tabela, RangeDados: Variant;
begin
  planilha := Excel.ActiveWorkbook.ActiveSheet;
  RangeDados := planilha.Range[planilha.Cells[3, 1], planilha.Cells[LinhaFinal, ColunaFinal]];

  if planilha.ListObjects.Count > 0 then
    planilha.ListObjects.Item(1).Delete;

  Tabela := planilha.ListObjects.Add(1, RangeDados, False, 1, EmptyParam);
  Tabela.TableStyle := 'TableStyleLight2';
  Tabela.ShowAutoFilter := True;

  planilha.Columns.AutoFit;

  planilha.Cells[4, 1].Select;
  Excel.ActiveWindow.FreezePanes := True;
end;

end.
