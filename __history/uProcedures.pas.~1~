unit uProcedures;

interface

uses Vcl.Controls, Vcl.ComCtrls, Vcl.StdCtrls, System.SysUtils, System.Classes,
  Vcl.Dialogs, FireDAC.Comp.Client,data.DB;

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

implementation

uses
  udmDados;

procedure AlterarMetas(AQuery: TDataSet; const AAmplitude, ADourado: Double);
var
  LMetaAmarelo, LMetaVerde, LMetaDourado: Double;
begin
  // Validação básica para evitar erros de Access Violation
  if not Assigned(AQuery) then
    Exit;
  // Realiza os cálculos
  LMetaDourado := ADourado;
  LMetaVerde   := ADourado * AAmplitude;
  LMetaAmarelo := LMetaVerde * AAmplitude;
  // Garante que a Query está em modo de edição
  if not (AQuery.State in [dsEdit, dsInsert]) then
    AQuery.Edit;
  // Atribuição utilizando FieldByName para flexibilidade
  AQuery.FieldByName('META_AMARELO').AsFloat := LMetaAmarelo;
  AQuery.FieldByName('META_VERDE').AsFloat   := LMetaVerde;
  AQuery.FieldByName('META_DOURADO').AsFloat := LMetaDourado;
end;
procedure HabilitarBotaoPorQuery(AQuery: TFDQuery; AButton: TControl);
begin
  AButton.Enabled := not AQuery.IsEmpty;
end;

procedure textoProgressBar(barra: TProgressBar; texto: TLabel);
var
  percentual: Double;
begin
  // Calcula o percentual
  percentual := (barra.Position - barra.Min) / (barra.Max - barra.Min) * 100;

  // Atualiza o texto do Label
  texto.Caption := Format('%.0f%%', [percentual]);
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
  memo.Lines.Add('');
end;

procedure abrirEscala(linha, veiculo, tipo_dia, funcionario, validar: Integer;
  dataInicial, dataFinal: tDate; qry: TFDQuery);
begin
  with dmDados do
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
end;

procedure preencher_edit(consulta: string; edit: TEdit);
var
  sname: string;
  aux: Integer;
  Posicao: Integer;
begin
  Try
    If edit.Text <> '' then
    begin
      sname := dmDados.con.ExecSQLScalar(consulta, [edit.Text + '%']);
      If sname <> '' then
      begin
        Posicao := length(edit.Text);
        For aux := length(edit.Text) + 1 to length(sname) do
        begin
          edit.Text := edit.Text + sname[aux];
        end;
        edit.SelStart := Posicao;
        edit.SelLength := length(edit.Text);
      end;
    end;
  Except
  end;
end;

procedure formatarEdit(edit: TEdit);
begin
  edit.Text := FormatFloat('#.00', StrToFloat(edit.Text));
end;

procedure eventosTeclado(var Key: Word; Shift: TShiftState; campo: string);
begin
  if (ssAlt in Shift) and (chr(Key) in ['N', 'n']) then
  begin
    var
      NewString: string;
    var
      ClickedOK: Boolean;
    NewString := 'Digite senha.';
    ClickedOK := InputQuery('Input Box', 'Prompt', NewString);
    if NewString = '5421' then
    begin

      NewString := campo; // IntToStr(dmDados.qryParametrosSIC.AsInteger);
      ClickedOK := InputQuery('Input Box', 'Prompt', NewString);
      if ClickedOK then
      begin
        try
          if ((StrToInt(NewString) > 45578) and (StrToInt(NewString) < 46000))
          then
          begin
            dmDados.con.ExecSQL('update parametros set sic=:sic',
              [StrToInt(NewString)]);
          end
          else
            ShowMessage
              ('Valor fora do parâmetro, Valor mínimo 45600 e Valor máximo 46000');
        except
          ShowMessage
            ('Valor fora do parâmetro, Valor mínimo 45600 e Valor máximo 46000');
        end;
      end;
    end
    else
      ShowMessage('Senha Inválida.');
  end;

end;

end.
