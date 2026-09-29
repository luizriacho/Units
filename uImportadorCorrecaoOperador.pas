unit uImportadorCorrecaoOperador;

interface

uses
  System.SysUtils,
  System.Classes,
  System.Variants,
  Winapi.ActiveX,
  ComObj,
  Vcl.Dialogs,
  Vcl.Forms,
  FireDAC.Comp.Client,
  FireDAC.Stan.Param,
  FireDAC.Stan.Option,
  FireDAC.Stan.Error,
  FireDAC.DatS,
  FireDAC.Phys.Intf,
  FireDAC.DApt.Intf,
  FireDAC.DApt;

type
  TImportadorCorrecaoOperador = class
  public
    /// <summary>
    /// Processa a planilha de correções de operador e atualiza a MOVIMENTO_TESOURARIA.
    /// </summary>
    /// <param name="AConnection">Conexão FireDAC ativa com o Firebird</param>
    /// <param name="ACaminhoArquivo">Caminho do arquivo .xls ou .xlsx</param>
    /// <param name="ATotalProcessados">Retorna o total de linhas lidas com dados válidos</param>
    /// <param name="ATotalAtualizados">Retorna o total de registros atualizados no banco</param>
    /// <param name="ATotalNaoEncontrados">Retorna o total de matrículas não localizadas na tabela FUNCIONARIO</param>
    /// <param name="AMensagemErro">Retorna a mensagem em caso de falha</param>
    class function ProcessarPlanilha(
      AConnection: TFDConnection;
      const ACaminhoArquivo: string;
      out ATotalProcessados: Integer;
      out ATotalAtualizados: Integer;
      out ATotalNaoEncontrados: Integer;
      out AMensagemErro: string
    ): Boolean;
  end;

implementation

class function TImportadorCorrecaoOperador.ProcessarPlanilha(
  AConnection: TFDConnection;
  const ACaminhoArquivo: string;
  out ATotalProcessados: Integer;
  out ATotalAtualizados: Integer;
  out ATotalNaoEncontrados: Integer;
  out AMensagemErro: string
): Boolean;
var
  vExcel: Variant;
  vWorkbook: Variant;
  vSheet: Variant;
  vQueryBuscaFun: TFDQuery;
  vQueryUpdateMov: TFDQuery;
  iLinha: Integer;
  iTotalLinhas: Integer;
  vValueId: Variant;
  vValueMatricula: Variant;
  iIdMovimento: Integer;
  iChapa: Integer;
  iChaveFun: Integer;
begin
  Result := False;
  AMensagemErro := '';
  ATotalProcessados := 0;
  ATotalAtualizados := 0;
  ATotalNaoEncontrados := 0;

  if not FileExists(ACaminhoArquivo) then
  begin
    AMensagemErro := 'Arquivo não encontrado: ' + ACaminhoArquivo;
    Exit;
  end;

  if not Assigned(AConnection) or not AConnection.Connected then
  begin
    AMensagemErro := 'Conexão com o banco de dados não está ativa.';
    Exit;
  end;

  CoInitialize(nil);
  vExcel := Unassigned;
  vWorkbook := Unassigned;
  vSheet := Unassigned;

  vQueryBuscaFun := TFDQuery.Create(nil);
  vQueryUpdateMov := TFDQuery.Create(nil);

  try
    try
      // Configura queries
      vQueryBuscaFun.Connection := AConnection;
      vQueryBuscaFun.SQL.Text :=
        'SELECT ' + sLineBreak +
        '    chave_fun ' + sLineBreak +
        '  , nome ' + sLineBreak +
        'FROM funcionario ' + sLineBreak +
        'WHERE chapa = :chapa';

      vQueryUpdateMov.Connection := AConnection;
      vQueryUpdateMov.SQL.Text :=
        'UPDATE movimento_tesouraria ' + sLineBreak +
        'SET chave_fun = :chave_fun ' + sLineBreak +
        'WHERE id_movimento_tesouraria = :id_movimento_tesouraria';

      // Inicializa instância do Excel via OLE
      try
        vExcel := CreateOleObject('Excel.Application');
      except
        on E: Exception do
        begin
          AMensagemErro := 'Não foi possível inicializar o Excel via COM/OLE. ' + E.Message;
          Exit;
        end;
      end;

      vExcel.Visible := False;
      vExcel.DisplayAlerts := False;

      // Abre a planilha
      vWorkbook := vExcel.Workbooks.Open(ACaminhoArquivo);
      vSheet := vWorkbook.Worksheets[1];

      iTotalLinhas := vSheet.UsedRange.Rows.Count;

      // Inicia transação no Firebird
      if not AConnection.InTransaction then
        AConnection.StartTransaction;

      // Linha 4 é onde começam os dados (Linha 1 = Título, Linha 3 = Cabeçalho)
      for iLinha := 4 to iTotalLinhas do
      begin
        vValueId := vSheet.Cells[iLinha, 1].Value;        // Coluna A: Id
        vValueMatricula := vSheet.Cells[iLinha, 3].Value; // Coluna C: Matricula (Chapa)

        // Se o ID for nulo ou vazio, encerra ou pula
        if VarIsNull(vValueId) or VarIsEmpty(vValueId) or (Trim(VarToStr(vValueId)) = '') then
          Continue;

        iIdMovimento := StrToIntDef(VarToStr(vValueId), 0);
        iChapa := StrToIntDef(VarToStr(vValueMatricula), 0);

        if (iIdMovimento <= 0) or (iChapa <= 0) then
          Continue;

        Inc(ATotalProcessados);

        // Busca a CHAVE_FUN do funcionário pela CHAPA
        vQueryBuscaFun.Close;
        vQueryBuscaFun.ParamByName('chapa').AsInteger := iChapa;
        vQueryBuscaFun.Open;

        if not vQueryBuscaFun.IsEmpty then
        begin
          iChaveFun := vQueryBuscaFun.FieldByName('chave_fun').AsInteger;

          // Atualiza a chave_fun no movimento
          vQueryUpdateMov.Close;
          vQueryUpdateMov.ParamByName('chave_fun').AsInteger := iChaveFun;
          vQueryUpdateMov.ParamByName('id_movimento_tesouraria').AsInteger := iIdMovimento;
          vQueryUpdateMov.ExecSQL;

          Inc(ATotalAtualizados);
        end
        else
        begin
          Inc(ATotalNaoEncontrados);
        end;
      end;

      // Confirma as alterações no banco de dados
      if AConnection.InTransaction then
        AConnection.Commit;

      Result := True;

    except
      on E: Exception do
      begin
        if AConnection.InTransaction then
          AConnection.Rollback;

        AMensagemErro := 'Erro ao processar planilha: ' + E.Message;
        Result := False;
      end;
    end;
  finally
    // Fecha a planilha e finaliza o Excel
    if not VarIsEmpty(vWorkbook) and not VarIsNull(vWorkbook) then
    begin
      vWorkbook.Close(False);
      vWorkbook := Unassigned;
    end;

    if not VarIsEmpty(vExcel) and not VarIsNull(vExcel) then
    begin
      vExcel.Quit;
      vExcel := Unassigned;
    end;

    vSheet := Unassigned;

    vQueryBuscaFun.Free;
    vQueryUpdateMov.Free;

    CoUninitialize;
  end;
end;

end.