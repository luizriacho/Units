unit UExcelService;

interface

uses
  System.SysUtils
, System.Variants
, Data.DB
, Winapi.ActiveX
, System.Win.ComObj
, Vcl.Dialogs
, System.Classes
, System.Generics.Collections;

type
  TExcelService = class
  private
    class function IsExcelInstalled: Boolean;
    class procedure FormatarPlanilhaExcel(const Excel: Variant;
      const LinhaFinal, ColunaFinal: Integer);
    class procedure AplicarEstiloTabela(const Sheet: Variant;
      LinhaInicial, LinhaFinal, ColunaFinal: Integer; Titulo: string);
    class procedure ExportarViaHTML(DataSet: TDataSet; ValorTag: Integer;
      const TituloCabecalho: string; Observacoes: TStrings = nil);
  public
    { Método Original: Mantido para compatibilidade com outros formulários do sistema }
    class procedure ExportarDataSet(DataSet: TDataSet; ValorTag: Integer;
      const TituloCabecalho: string; Observacoes: TStrings = nil);

    { Novo Método: Mestre/Detalhe com Hyperlinks para Ocorrências }
    class procedure ExportarMestreDetalhe(DataSetMestre, DataSetDetalhe
      : TDataSet; TagMestre, TagDetalhe: Integer; const TituloMestre: string;
      CampoLink: string);
  end;

implementation

uses
  Winapi.Windows
, Winapi.ShellAPI;

{ TExcelService }

class function TExcelService.IsExcelInstalled: Boolean;
var
  ClassID: TCLSID;
begin
  Result := CLSIDFromProgID('Excel.Application', ClassID) = S_OK;
end;

class procedure TExcelService.ExportarDataSet(DataSet: TDataSet;
  ValorTag: Integer; const TituloCabecalho: string;
  Observacoes: TStrings = nil);
var
  Linha, coluna, ColExcel: Integer;
  planilha, Sheet, Dados, vRange: Variant;
  UltimaColunaPreenchida: Integer;
  TotalRegistros: Integer;
  PathImages, LogoCliente, LogoSic: string;
  PosicaoEsquerdaSic: Double;
  LinhaRodape, i: Integer;
const
  xlCenter = -4108;

  // Função auxiliar para converter campo em Variant seguro
  function FieldToVariant(Field: TField): Variant;
  begin
    if Field.IsNull then
    begin
      Result := '';
      Exit;
    end;
    case Field.DataType of
      ftBCD, ftFMTBcd:
        Result := Field.AsFloat;
      ftCurrency:
        Result := Field.AsCurrency;
      ftInteger, ftSmallint, ftWord, ftLargeint, ftAutoInc:
        Result := Field.AsInteger;
      ftFloat, ftExtended, ftSingle:
        Result := Field.AsFloat;
      ftDate, ftDateTime, ftTime, ftTimeStamp:
        Result := Field.AsDateTime;
      ftBoolean:
        Result := Field.AsBoolean;
    else
      Result := Field.AsString;
    end;
  end;

begin
  if not Assigned(DataSet) or (DataSet.IsEmpty) then
    Exit;

  // Se o Excel não estiver instalado, faz o Fallback gerando um arquivo limpo
  if not IsExcelInstalled then
  begin
    ExportarViaHTML(DataSet, ValorTag, TituloCabecalho, Observacoes);
    Exit;
  end;

  PathImages := ExtractFilePath(ParamStr(0)) + 'images\';
  LogoCliente := PathImages + 'logoCliente.png';
  LogoSic     := PathImages + 'logoSic.png';

  DataSet.Last;
  TotalRegistros := DataSet.RecordCount;
  DataSet.First;

  DataSet.DisableControls;
  try
    planilha := CreateOleObject('Excel.Application');
    planilha.WorkBooks.Add;
    planilha.Visible := True;
    planilha.ScreenUpdating := False;
    Sheet := planilha.ActiveWorkbook.ActiveSheet;

    UltimaColunaPreenchida := 0;
    for coluna := 0 to DataSet.FieldCount - 1 do
      if DataSet.Fields[coluna].Tag = ValorTag then
        Inc(UltimaColunaPreenchida);

    if UltimaColunaPreenchida = 0 then
      Exit;

    Sheet.Rows.Item[1].RowHeight := 25;
    Sheet.Rows.Item[2].RowHeight := 25;

    Dados := VarArrayCreate([1, TotalRegistros + 1, 1, UltimaColunaPreenchida],
      varVariant);

    // Cabeçalho
    ColExcel := 1;
    for coluna := 0 to DataSet.FieldCount - 1 do
      if DataSet.Fields[coluna].Tag = ValorTag then
      begin
        Dados[1, ColExcel] := DataSet.Fields[coluna].DisplayLabel;
        Inc(ColExcel);
      end;

    // Dados
    Linha := 2;
    while not DataSet.Eof do
    begin
      ColExcel := 1;
      for coluna := 0 to DataSet.FieldCount - 1 do
        if DataSet.Fields[coluna].Tag = ValorTag then
        begin
          Dados[Linha, ColExcel] := FieldToVariant(DataSet.Fields[coluna]);
          Inc(ColExcel);
        end;
      Inc(Linha);
      DataSet.Next;
    end;

    vRange := Sheet.Range[Sheet.Cells[3, 1],
      Sheet.Cells[2 + Linha - 1, UltimaColunaPreenchida]];
    vRange.Value := Dados;

    FormatarPlanilhaExcel(planilha, 2 + Linha - 1, UltimaColunaPreenchida);

    // Título / cabeçalho mesclado
    vRange := Sheet.Range[Sheet.Cells[1, 1],
      Sheet.Cells[2, UltimaColunaPreenchida]];
    vRange.Merge;
    vRange.Value := TituloCabecalho;
    vRange.Font.Bold := True;
    vRange.Font.Size := 14;
    vRange.HorizontalAlignment := xlCenter;
    vRange.VerticalAlignment   := xlCenter;

    // Observações no rodapé
    if Assigned(Observacoes) and (Observacoes.Count > 0) then
    begin
      LinhaRodape := 3 + TotalRegistros + 2;
      for i := 0 to Observacoes.Count - 1 do
      begin
        vRange := Sheet.Cells[LinhaRodape + i, 1];
        vRange.Value      := Observacoes[i];
        vRange.Font.Italic := True;
        vRange.Font.Color  := $00555555;
      end;
    end;

    // Logos
    if FileExists(LogoCliente) then
      Sheet.Shapes.AddPicture(LogoCliente, False, True, 5, 5, 65, 38);

    if FileExists(LogoSic) then
    begin
      PosicaoEsquerdaSic :=
        Sheet.Cells[1, UltimaColunaPreenchida].Left +
        Sheet.Cells[1, UltimaColunaPreenchida].Width - 65;
      Sheet.Shapes.AddPicture(LogoSic, False, True,
        PosicaoEsquerdaSic, 5, 42, 42);
    end;

  finally
    DataSet.EnableControls;
    if VarType(planilha) = varDispatch then
      planilha.ScreenUpdating := True;
  end;
end;

class procedure TExcelService.ExportarViaHTML(DataSet: TDataSet; ValorTag: Integer;
  const TituloCabecalho: string; Observacoes: TStrings = nil);
var
  ArquivoTXT: TStringList;
  TempFile: string;
  coluna: Integer;
begin
  ArquivoTXT := TStringList.Create;
  try
    ArquivoTXT.Add('<html>');
    ArquivoTXT.Add('<head><meta charset="utf-8"><style>');
    ArquivoTXT.Add('body { font-family: Arial, sans-serif; }');
    ArquivoTXT.Add('table { border-collapse: collapse; width: 100%; }');
    ArquivoTXT.Add('th { background-color: #f2f2f2; color: #333; font-weight: bold; border: 1px solid #ddd; padding: 8px; }');
    ArquivoTXT.Add('td { border: 1px solid #ddd; padding: 8px; text-align: left; }');
    ArquivoTXT.Add('h2 { color: #002060; }');
    ArquivoTXT.Add('</style></head><body>');

    ArquivoTXT.Add('<h2>' + TituloCabecalho + '</h2>');
    ArquivoTXT.Add('<table><thead><tr>');

    for coluna := 0 to DataSet.FieldCount - 1 do
      if DataSet.Fields[coluna].Tag = ValorTag then
        ArquivoTXT.Add('<th>' + DataSet.Fields[coluna].DisplayLabel + '</th>');

    ArquivoTXT.Add('</tr></thead><tbody>');

    DataSet.First;
    while not DataSet.Eof do
    begin
      ArquivoTXT.Add('<tr>');
      for coluna := 0 to DataSet.FieldCount - 1 do
        if DataSet.Fields[coluna].Tag = ValorTag then
          ArquivoTXT.Add('<td>' + DataSet.Fields[coluna].DisplayText + '</td>');
      ArquivoTXT.Add('</tr>');
      DataSet.Next;
    end;
    ArquivoTXT.Add('</tbody></table>');

    if Assigned(Observacoes) and (Observacoes.Count > 0) then
    begin
      ArquivoTXT.Add('<br><br><i>');
      for coluna := 0 to Observacoes.Count - 1 do
        ArquivoTXT.Add('<p>' + Observacoes[coluna] + '</p>');
      ArquivoTXT.Add('</i>');
    end;

    ArquivoTXT.Add('</body></html>');

    TempFile := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP')) +
                'Exportacao_' + FormatDateTime('hhmmss', Now) + '.xls';
    ArquivoTXT.SaveToFile(TempFile, TEncoding.UTF8);

    // ShellExecute nativo com Winapi: abre com a aplicação padrão associada ao .xls
    ShellExecute(0, 'open', PChar(TempFile), nil, nil, SW_SHOWNORMAL);
  finally
    ArquivoTXT.Free;
  end;
end;

class procedure TExcelService.ExportarMestreDetalhe(DataSetMestre, DataSetDetalhe: TDataSet;
  TagMestre, TagDetalhe: Integer; const TituloMestre: string; CampoLink: string);
var
  Excel, Workbook, SheetResumo, SheetDetalhes, vRange: Variant;
  Linha, Col, ColExcel, UltimaColMestre, UltimaColDetalhe: Integer;
  DadosMestre, DadosDetalhe: Variant;
  DicPosicao: TDictionary<string, Integer>;
  i, ColTotalOco: Integer;
  PathImages, LogoCliente, LogoSic: string;
  PosicaoEsquerdaSic: Double;
begin
  if not IsExcelInstalled then
  begin
    ShowMessage('A exportação mestre/detalhe com recursos de hiperlink dinâmico exige o Microsoft Excel instalado.');
    Exit;
  end;

  PathImages  := ExtractFilePath(ParamStr(0)) + 'images\';
  LogoCliente := PathImages + 'logoCliente.png';
  LogoSic     := PathImages + 'logoSic.png';

  Excel := CreateOleObject('Excel.Application');
  Excel.Visible := True;
  Workbook := Excel.Workbooks.Add;
  DicPosicao := TDictionary<string, Integer>.Create;

  DataSetMestre.DisableControls;
  DataSetDetalhe.DisableControls;
  try
    SheetResumo := Workbook.Worksheets.Item[1];
    SheetResumo.Name := 'Resumo Operadores';
    SheetResumo.Rows.Item[1].RowHeight := 25;
    SheetResumo.Rows.Item[2].RowHeight := 25;

    SheetDetalhes := Workbook.Worksheets.Add(EmptyParam, SheetResumo);
    SheetDetalhes.Name := 'Ocorrencias Detalhadas';

    UltimaColDetalhe := 0;
    for Col := 0 to DataSetDetalhe.FieldCount - 1 do
      if DataSetDetalhe.Fields[Col].Tag = TagDetalhe then Inc(UltimaColDetalhe);

    DataSetDetalhe.Last;
    DadosDetalhe := VarArrayCreate([1, DataSetDetalhe.RecordCount + 1, 1, UltimaColDetalhe], varVariant);

    ColExcel := 1;
    for Col := 0 to DataSetDetalhe.FieldCount - 1 do
      if DataSetDetalhe.Fields[Col].Tag = TagDetalhe then
      begin
        DadosDetalhe[1, ColExcel] := DataSetDetalhe.Fields[Col].DisplayLabel;
        Inc(ColExcel);
      end;

    DataSetDetalhe.First;
    Linha := 2;
    while not DataSetDetalhe.Eof do
    begin
      if not DicPosicao.ContainsKey(DataSetDetalhe.FieldByName(CampoLink).AsString) then
        DicPosicao.Add(DataSetDetalhe.FieldByName(CampoLink).AsString, Linha + 2);

      ColExcel := 1;
      for Col := 0 to DataSetDetalhe.FieldCount - 1 do
        if DataSetDetalhe.Fields[Col].Tag = TagDetalhe then
        begin
          DadosDetalhe[Linha, ColExcel] := DataSetDetalhe.Fields[Col].Value;
          Inc(ColExcel);
        end;
      Inc(Linha);
      DataSetDetalhe.Next;
    end;

    vRange := SheetDetalhes.Range[SheetDetalhes.Cells[3, 1], SheetDetalhes.Cells[1 + Linha, UltimaColDetalhe]];
    vRange.Value := DadosDetalhe;
    AplicarEstiloTabela(SheetDetalhes, 3, 1 + Linha, UltimaColDetalhe, 'DETALHAMENTO DE OCORRÊNCIAS');

    UltimaColMestre := 0;
    ColTotalOco := 0;
    ColExcel := 1;
    for Col := 0 to DataSetMestre.FieldCount - 1 do
      if DataSetMestre.Fields[Col].Tag = TagMestre then
      begin
        Inc(UltimaColMestre);
        if DataSetMestre.Fields[Col].FieldName = 'TOTALOCO' then ColTotalOco := ColExcel;
        Inc(ColExcel);
      end;

    DataSetMestre.Last;
    DadosMestre := VarArrayCreate([1, DataSetMestre.RecordCount + 1, 1, UltimaColMestre], varVariant);

    ColExcel := 1;
    for Col := 0 to DataSetMestre.FieldCount - 1 do
      if DataSetMestre.Fields[Col].Tag = TagMestre then
      begin
        DadosMestre[1, ColExcel] := DataSetMestre.Fields[Col].DisplayLabel;
        Inc(ColExcel);
      end;

    DataSetMestre.First;
    Linha := 2;
    while not DataSetMestre.Eof do
    begin
      ColExcel := 1;
      for Col := 0 to DataSetMestre.FieldCount - 1 do
        if DataSetMestre.Fields[Col].Tag = TagMestre then
        begin
          DadosMestre[Linha, ColExcel] := DataSetMestre.Fields[Col].Value;
          Inc(ColExcel);
        end;
      Inc(Linha);
      DataSetMestre.Next;
    end;

    vRange := SheetResumo.Range[SheetResumo.Cells[3, 1], SheetResumo.Cells[1 + Linha, UltimaColMestre]];
    vRange.Value := DadosMestre;
    AplicarEstiloTabela(SheetResumo, 3, 1 + Linha, UltimaColMestre, TituloMestre);

    if FileExists(LogoCliente) then
      SheetResumo.Shapes.AddPicture(LogoCliente, False, True, 5, 5, 65, 38);

    if FileExists(LogoSic) then
    begin
      PosicaoEsquerdaSic := SheetResumo.Cells[1, UltimaColMestre].Left +
                            SheetResumo.Cells[1, UltimaColMestre].Width - 65;
      SheetResumo.Shapes.AddPicture(LogoSic, False, True, PosicaoEsquerdaSic, 5, 42, 42);
    end;

    for i := 4 to (3 + DataSetMestre.RecordCount) do
    begin
      var ChapaLink := VarToStr(SheetResumo.Cells[i, 2].Value);
      if DicPosicao.ContainsKey(ChapaLink) and (ColTotalOco > 0) then
      begin
        SheetResumo.Hyperlinks.Add(
          SheetResumo.Cells[i, ColTotalOco],
          '',
          '''' + SheetDetalhes.Name + '''!A' + IntToStr(DicPosicao[ChapaLink]),
          'Ver Detalhes'
        );
        SheetResumo.Cells[i, ColTotalOco].Font.Color := $0000FF;
        SheetResumo.Cells[i, ColTotalOco].Font.Bold := True;
      end;
    end;

    SheetResumo.Activate;

  finally
    DicPosicao.Free;
    DataSetMestre.EnableControls;
    DataSetDetalhe.EnableControls;
  end;
end;

class procedure TExcelService.AplicarEstiloTabela(const Sheet: Variant;
  LinhaInicial, LinhaFinal, ColunaFinal: Integer; Titulo: string);
var
  vRange, Tabela: Variant;
begin
  vRange := Sheet.Range[Sheet.Cells[1, 1], Sheet.Cells[2, ColunaFinal]];
  vRange.Merge;
  vRange.Value := Titulo;
  vRange.Font.Bold := True;
  vRange.Font.Size := 14;
  vRange.HorizontalAlignment := -4108;

  vRange := Sheet.Range[Sheet.Cells[LinhaInicial, 1],
    Sheet.Cells[LinhaFinal, ColunaFinal]];
  Tabela := Sheet.ListObjects.Add(1, vRange, False, 1);
  Tabela.TableStyle := 'TableStyleLight2';
  Sheet.Columns.AutoFit;
end;

class procedure TExcelService.FormatarPlanilhaExcel(const Excel: Variant;
  const LinhaFinal, ColunaFinal: Integer);
var
  planilha, Tabela, RangeDados: Variant;
begin
  planilha := Excel.ActiveWorkbook.ActiveSheet;
  RangeDados := planilha.Range[planilha.Cells[3, 1], planilha.Cells[LinhaFinal,
    ColunaFinal]];
  if planilha.ListObjects.Count > 0 then
    planilha.ListObjects.Item(1).Delete;
  Tabela := planilha.ListObjects.Add(1, RangeDados, False, 1);
  Tabela.TableStyle := 'TableStyleLight2';
  planilha.Columns.AutoFit;
  planilha.Cells[4, 1].Select;
  Excel.ActiveWindow.FreezePanes := True;
end;

end.
