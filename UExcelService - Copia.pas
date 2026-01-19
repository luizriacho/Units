unit UExcelService;

interface

uses
  System.SysUtils, System.Variants, Data.DB, Winapi.ActiveX, System.Win.ComObj;

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
  planilha, Sheet, Dados: Variant;
  UltimaColunaPreenchida: Integer;
  TotalRegistros: Integer;
begin
  if not Assigned(DataSet) or (DataSet.IsEmpty) then
    Exit;

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

    // 2. TÍTULO (Obrigatório ser antes do Range de dados)
    Sheet.Range[Sheet.Cells[1, 1], Sheet.Cells[2, UltimaColunaPreenchida]].Merge;
    Sheet.Cells[1, 1].Value := TituloCabecalho;
    Sheet.Cells[1, 1].Font.Bold := True;
    Sheet.Cells[1, 1].Font.Size := 14;
    Sheet.Cells[1, 1].HorizontalAlignment := -4108; // xlCenter

    // 3. Preparar Array (Header + Records)
    // O tamanho deve ser TotalRegistros + 1 (para o Header)
    Dados := VarArrayCreate([1, TotalRegistros + 1, 1, UltimaColunaPreenchida], varVariant);

    // Header da Tabela (Linha 1 do Array)
    ColExcel := 1;
    for coluna := 0 to DataSet.FieldCount - 1 do
      if DataSet.Fields[coluna].Tag = ValorTag then
      begin
        Dados[1, ColExcel] := DataSet.Fields[coluna].DisplayLabel;
        Inc(ColExcel);
      end;

    // Dados (Inicia na Linha 2 do Array)
    Linha := 2;
    DataSet.First;
    while not DataSet.Eof do
    begin
      ColExcel := 1;
      for coluna := 0 to DataSet.FieldCount - 1 do
        if DataSet.Fields[coluna].Tag = ValorTag then
        begin
          try
            if DataSet.Fields[coluna].IsNull then
              Dados[Linha, ColExcel] := ''
            else if DataSet.Fields[coluna] is TNumericField then
              Dados[Linha, ColExcel] := DataSet.Fields[coluna].AsFloat // Resolve erro de tipo
            else
              Dados[Linha, ColExcel] := DataSet.Fields[coluna].Value;
          except
            Dados[Linha, ColExcel] := DataSet.Fields[coluna].AsString;
          end;
          Inc(ColExcel);
        end;
      Inc(Linha);
      DataSet.Next;
    end;

    // 4. DESPEJA O ARRAY (Começa na Linha 3 do Excel)
    // O Range final deve ir até (3 + Total de Linhas do Array - 1)
    Sheet.Range[Sheet.Cells[3, 1], Sheet.Cells[3 + TotalRegistros, UltimaColunaPreenchida]].Value := Dados;

    // 5. FORMATAÇÃO (Passando o número exato de linhas totais: Header + Dados)
    FormatarPlanilhaExcel(planilha, 3 + TotalRegistros, UltimaColunaPreenchida);

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
