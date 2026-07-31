unit uLogUtils;

interface

uses
  System.SysUtils
, System.Classes
, System.TypInfo
, Vcl.Controls
, Vcl.Forms
, FireDAC.Comp.Client;

type
  TLogUtils = class
  public
    class procedure RegistrarLog(
      AConnection: TFDConnection;
      const AUsuario: string;
      const ATipo: string;
      const AAcao: string = '';
      const AFormulario: string = ''
    ); overload;

    // Sobrecarga conveniente que extrai Form e Caption do componente automaticamente
    class procedure RegistrarLogComponente(
      AConnection: TFDConnection;
      const AUsuario: string;
      const ATipo: string;
      AComponente: TComponent
    ); overload;
  end;

implementation

class procedure TLogUtils.RegistrarLog(
  AConnection: TFDConnection;
  const AUsuario: string;
  const ATipo: string;
  const AAcao: string;
  const AFormulario: string
);
var
  Qry: TFDQuery;
begin
  if not Assigned(AConnection) or not AConnection.Connected then
    Exit;

  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := AConnection;
    Qry.SQL.Text := 
      'INSERT INTO LOG_ACESSO (' + sLineBreak +
      '    DATA_HORA' + sLineBreak +
      '  , USUARIO' + sLineBreak +
      '  , TIPO' + sLineBreak +
      '  , FORMULARIO' + sLineBreak +
      '  , ACAO' + sLineBreak +
      ') VALUES (' + sLineBreak +
      '    CURRENT_TIMESTAMP' + sLineBreak +
      '  , :USUARIO' + sLineBreak +
      '  , :TIPO' + sLineBreak +
      '  , :FORMULARIO' + sLineBreak +
      '  , :ACAO' + sLineBreak +
      ')';

    Qry.ParamByName('USUARIO').AsString := Copy(Trim(AUsuario), 1, 30);
    Qry.ParamByName('TIPO').AsString := Copy(Trim(ATipo), 1, 20);
    
    if AFormulario.IsEmpty then
      Qry.ParamByName('FORMULARIO').Clear
    else
      Qry.ParamByName('FORMULARIO').AsString := Copy(Trim(AFormulario), 1, 50);

    if AAcao.IsEmpty then
      Qry.ParamByName('ACAO').Clear
    else
      Qry.ParamByName('ACAO').AsString := Copy(Trim(AAcao), 1, 100);

    Qry.ExecSQL;
  finally
    Qry.Free;
  end;
end;

class procedure TLogUtils.RegistrarLogComponente(
  AConnection: TFDConnection;
  const AUsuario: string;
  const ATipo: string;
  AComponente: TComponent
);
var
  NomeForm: string;
  DescricaoAcao: string;
  FormPai: TCustomForm;
begin
  NomeForm := '';
  DescricaoAcao := '';

  if Assigned(AComponente) then
  begin
    // Identifica o Form pai do componente
    if AComponente is TControl then
    begin
      FormPai := GetParentForm(TControl(AComponente));
      if Assigned(FormPai) then
        NomeForm := FormPai.Name;
    end
    else if AComponente.Owner is TCustomForm then
    begin
      NomeForm := TCustomForm(AComponente.Owner).Name;
    end;

    // Tenta pegar a legenda/caption ou o nome do componente
    if IsPublishedProp(AComponente, 'Caption') then
      DescricaoAcao := GetPropValue(AComponente, 'Caption')
    else if IsPublishedProp(AComponente, 'Text') then
      DescricaoAcao := GetPropValue(AComponente, 'Text')
    else
      DescricaoAcao := AComponente.Name;

    // Adiciona o nome do componente para ficar claro (ex: "btnSalvar - Salvar Dados")
    if (DescricaoAcao <> AComponente.Name) and (AComponente.Name <> '') then
      DescricaoAcao := AComponente.Name + ' - ' + DescricaoAcao;
  end;

  RegistrarLog(AConnection, AUsuario, ATipo, DescricaoAcao, NomeForm);
end;

end.