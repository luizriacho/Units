unit uProcedures;

interface

uses
  Vcl.Controls, Vcl.ComCtrls, Vcl.StdCtrls, System.SysUtils, System.Classes,
  Vcl.Dialogs, FireDAC.Comp.Client, data.DB, System.DateUtils,
  JvToolEdit, Vcl.DBGrids, Vcl.ExtCtrls, Vcl.Buttons; // Unidades Vcl.ExtCtrls e Vcl.Buttons adicionadas

type
  { Enumeração para a direção da navegação }
  TDirecaoMes = (dmAnterior, dmProximo);

  { Protótipos das procedures }
procedure textoProgressBar(barra: TProgressBar; texto: TLabel);
procedure abrirEscala(linha, veiculo, tipo_dia, funcionario, validar: Integer;
  dataInicial, dataFinal: tDate; qry: TFDQuery);
procedure preencher_edit(consulta: string; edit: TEdit);
procedure formatarEdit(edit: TEdit);
procedure eventosTeclado(var Key: Word; Shift: TShiftState; campo: string);
procedure iniciarMemo(memo: TMemo);
procedure finalizarMemo(memo: TMemo; contador: Integer; texto: string);
procedure AlterarMetas(AQuery: TDataSet; const AAmplitude, ADourado: Double);
procedure HabilitarBotaoPorQuery(AQuery: TFDQuery; AButton: TControl);
procedure OrdenarGrid(Column: TColumn);
procedure ExtrairDatasRelatorio(const texto: string;
  out dtInicio, dtFim: TDateTime);
procedure ExtrairDatasRelatorioTelemetria(const ATexto: string;
  out dtInicio, dtFim: TDateTime);
{ MÉTODO GENÉRICO PARA JVDateEdit }
procedure NavegarMes(ADirecao: TDirecaoMes;
  const ADtInicio, ADtFim: TJvDateEdit);
{ MÉTODO GENÉRICO PARA ALTERNAR PAINEL }
procedure AlternarPainelGenerico(
  APainel: TPanel;
  ABtnToggle: TControl;
  var AFlagEstado: Boolean;
  AExibir: Boolean;
  ATamanhoExpandido: Integer;
  const AControlesExtras: array of TControl
);

implementation

uses
  udmDados;

procedure ExtrairDatasRelatorioTelemetria(const ATexto: string;
  out dtInicio, dtFim: TDateTime);
var
  posDE, paraPara: Integer;
  sIni, sFim: string;
  fs: TFormatSettings;

  function ConverterDataHora(const s: string): TDateTime;
  var
    sData, sHora: string;
    partes: TArray<string>;
  begin
    partes := s.Split([' ']);
    sData := partes[0];
    if Length(partes) > 1 then
      sHora := partes[1]
    else
      sHora := '00:00';
    Result := StrToDate(sData, fs) + StrToTime(sHora, fs);
  end;

begin
  fs := TFormatSettings.Create('pt-BR');
  fs.ShortDateFormat := 'dd/mm/yyyy';
  fs.ShortTimeFormat := 'hh:mm';
  fs.TimeSeparator := ':';
  fs.DateSeparator := '/';

  posDE := Pos('De ', ATexto);
  paraPara := Pos('Para', ATexto);

  sIni := Trim(Copy(ATexto, posDE + 3, paraPara - (posDE + 3)));
  sFim := Trim(Copy(ATexto, paraPara + 4, MaxInt));
  sIni := Trim(Copy(sIni, 1, 16));
  sFim := Trim(Copy(sFim, 1, 16));

  dtInicio := ConverterDataHora(sIni);
  dtFim := ConverterDataHora(sFim);
end;

procedure ExtrairDatasRelatorio(const texto: string;
  out dtInicio, dtFim: TDateTime);
var
  posDE, paraPara: Integer;
  sIni, sFim: string;
  fs: TFormatSettings;

  function ConverterDataHora(const s: string): TDateTime;
  var
    sData, sHora: string;
    partes: TArray<string>;
  begin
    partes := s.Split([' ']);
    sData := partes[0]; // dd/mm/yyyy
    if Length(partes) > 1 then
      sHora := partes[1] // hh:mm
    else
      sHora := '00:00';

    Result := StrToDate(sData, fs) + StrToTime(sHora, fs);
  end;

begin
  fs := TFormatSettings.Create('pt-BR');
  fs.ShortDateFormat := 'dd/mm/yyyy';
  fs.ShortTimeFormat := 'hh:mm';
  fs.TimeSeparator := ':';
  fs.DateSeparator := '/';

  posDE := Pos('De ', texto);
  paraPara := Pos('Para', texto);

  sIni := Trim(Copy(texto, posDE + 3, paraPara - (posDE + 3)));
  sFim := Trim(Copy(texto, paraPara + 4, MaxInt));
  sIni := Trim(Copy(sIni, 1, 16));
  sFim := Trim(Copy(sFim, 1, 16));

  dtInicio := ConverterDataHora(sIni);
  dtFim := ConverterDataHora(sFim);
end;

procedure NavegarMes(ADirecao: TDirecaoMes;
  const ADtInicio, ADtFim: TJvDateEdit);
var
  LData: tDate;
begin
  // Extrai a data do componente Jv
  LData := ADtInicio.Date;

  if ADirecao = dmAnterior then
    LData := IncMonth(LData, -1)
  else
    LData := IncMonth(LData, 1);

  // Define o período completo do mês
  ADtInicio.Date := StartOfAMonth(YearOf(LData), MonthOf(LData));
  ADtFim.Date := EndOfAMonth(YearOf(LData), MonthOf(LData));
end;

procedure AlterarMetas(AQuery: TDataSet; const AAmplitude, ADourado: Double);
var
  LMetaAmarelo, LMetaVerde, LMetaDourado: Double;
begin
  if not Assigned(AQuery) then
    Exit;
  LMetaDourado := ADourado;
  LMetaVerde := ADourado * AAmplitude;
  LMetaAmarelo := LMetaVerde * AAmplitude;
  if not(AQuery.State in [dsEdit, dsInsert]) then
    AQuery.edit;
  AQuery.FieldByName('META_AMARELO').AsFloat := LMetaAmarelo;
  AQuery.FieldByName('META_VERDE').AsFloat := LMetaVerde;
  AQuery.FieldByName('META_DOURADO').AsFloat := LMetaDourado;
end;

procedure HabilitarBotaoPorQuery(AQuery: TFDQuery; AButton: TControl);
begin
  AButton.Enabled := not AQuery.IsEmpty;
end;

procedure textoProgressBar(barra: TProgressBar; texto: TLabel);
begin
  if (barra.Max - barra.Min) > 0 then
    texto.Caption := Format('%.0f%%',
      [(barra.Position - barra.Min) / (barra.Max - barra.Min) * 100]);
end;

procedure iniciarMemo(memo: TMemo);
begin
  memo.Visible := true;
  memo.Height := 200;
  memo.Lines.Add(UpperCase('hora início ' + (formatDateTime('dd|mmm|yy', now)) +
    '   ' + formatDateTime('ttt', time)));
end;

procedure finalizarMemo(memo: TMemo; contador: Integer; texto: string);
begin
  memo.Lines.Add(texto + IntToStr(contador));
  memo.Lines.Add(UpperCase('hora final ' + (formatDateTime('dd|mmm|yy', now)) +
    '   ' + formatDateTime('ttt', time)));
  memo.Lines.Add('');
end;

procedure abrirEscala(linha, veiculo, tipo_dia, funcionario, validar: Integer;
  dataInicial, dataFinal: tDate; qry: TFDQuery);
begin
  qry.Close;
  if linha = 0 then
    qry.Params[0].Clear
  else
    qry.Params[0].Value := linha;
  if veiculo = 0 then
    qry.Params[1].Clear
  else
    qry.Params[1].Value := veiculo;
  if tipo_dia = 0 then
    qry.Params[2].Clear
  else
    qry.Params[2].Value := tipo_dia;
  if funcionario = 0 then
    qry.Params[3].Clear
  else
    qry.Params[3].Value := funcionario;
  qry.Params[4].AsDate := dataInicial;
  qry.Params[5].AsDate := dataFinal;
  qry.Params[6].Value := validar;
  qry.Open;
end;

procedure preencher_edit(consulta: string; edit: TEdit);
var
  sname: string;
  aux, Posicao: Integer;
begin
  Try
    If edit.Text <> '' then
    begin
      sname := dmDados.con.ExecSQLScalar(consulta, [edit.Text + '%']);
      If sname <> '' then
      begin
        Posicao := Length(edit.Text);
        For aux := Length(edit.Text) + 1 to Length(sname) do
          edit.Text := edit.Text + sname[aux];
        edit.SelStart := Posicao;
        edit.SelLength := Length(edit.Text);
      end;
    end;
  Except
  end;
end;

procedure formatarEdit(edit: TEdit);
begin
  edit.Text := FormatFloat('#.00', StrToFloatDef(edit.Text, 0));
end;

procedure eventosTeclado(var Key: Word; Shift: TShiftState; campo: string);
begin
  if (ssAlt in Shift) and (chr(Key) in ['N', 'n']) then
  begin
    var
      NewString: string := 'Digite senha.';
    if (InputQuery('Acesso', 'Senha:', NewString)) and (NewString = '5421') then
    begin
      NewString := campo;
      if InputQuery('Configuração', 'Valor:', NewString) then
      begin
        try
          dmDados.con.ExecSQL('update parametros set sic=:sic',
            [StrToInt(NewString)]);
        except
          ShowMessage('Erro ao atualizar');
        end;
      end;
    end;
  end;
end;

procedure OrdenarGrid(Column: TColumn);
var
  vDataset: TDataSet;
  vFieldName: string;
begin
  vDataset := Column.Field.DataSet;
  vFieldName := Column.FieldName;

  // Verificamos se o DataSet é um FDQuery (Firebird) para usar IndexFieldNames
  if vDataset is TFDQuery then
  begin
    try
      vDataset.DisableControls; // Melhora a performance visual
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

procedure AlternarPainelGenerico(
  APainel: TPanel;
  ABtnToggle: TControl;
  var AFlagEstado: Boolean;
  AExibir: Boolean;
  ATamanhoExpandido: Integer;
  const AControlesExtras: array of TControl
);
var
  I: Integer;
  vIsVertical: Boolean;
begin
  if not Assigned(APainel) then Exit;

  AFlagEstado := AExibir;
  vIsVertical := APainel.Align in [alTop, alBottom];

  APainel.DisableAlign;
  try
    if AFlagEstado then
    begin
      // Expande na dimensão correspondente
      if vIsVertical then
        APainel.Height := ATamanhoExpandido
      else
        APainel.Width := ATamanhoExpandido;

      // Ajusta o Caption do botão conforme a orientação
      if Assigned(ABtnToggle) then
      begin
        if vIsVertical then
        begin
          if ABtnToggle is TBitBtn then
            TBitBtn(ABtnToggle).Caption := '▲'
          else if ABtnToggle is TSpeedButton then
            TSpeedButton(ABtnToggle).Caption := '▲'
          else if ABtnToggle is TButton then
            TButton(ABtnToggle).Caption := '▲';
        end
        else
        begin
          if ABtnToggle is TBitBtn then
            TBitBtn(ABtnToggle).Caption := '◀ / ▶'
          else if ABtnToggle is TSpeedButton then
            TSpeedButton(ABtnToggle).Caption := '◀ / ▶'
          else if ABtnToggle is TButton then
            TButton(ABtnToggle).Caption := '◀ / ▶';
        end;
      end;
    end
    else
    begin
      // Recolhe na dimensão correspondente (mantém 28/30px para a barrinha do botão)
      if vIsVertical then
        APainel.Height := 28
      else
        APainel.Width := 30;

      // Ajusta o Caption do botão conforme a orientação
      if Assigned(ABtnToggle) then
      begin
        if vIsVertical then
        begin
          if ABtnToggle is TBitBtn then
            TBitBtn(ABtnToggle).Caption := '▼'
          else if ABtnToggle is TSpeedButton then
            TSpeedButton(ABtnToggle).Caption := '▼'
          else if ABtnToggle is TButton then
            TButton(ABtnToggle).Caption := '▼';
        end
        else
        begin
          if ABtnToggle is TBitBtn then
            TBitBtn(ABtnToggle).Caption := '▶'
          else if ABtnToggle is TSpeedButton then
            TSpeedButton(ABtnToggle).Caption := '▶'
          else if ABtnToggle is TButton then
            TButton(ABtnToggle).Caption := '▶';
        end;
      end;
    end;

    // Exibe ou oculta os controles internos ou painel filho passados
    for I := Low(AControlesExtras) to High(AControlesExtras) do
    begin
      if Assigned(AControlesExtras[I]) then
        AControlesExtras[I].Visible := AFlagEstado;
    end;
  finally
    APainel.EnableAlign;
  end;
end;
end.
