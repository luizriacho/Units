unit uControleAcesso;

interface

uses
  System.SysUtils, Vcl.Dialogs, FireDAC.Comp.Client;

type
  TPermissao = record
    CanAccess: Boolean;
    CanInsert: Boolean;
    CanEdit: Boolean;
    CanDelete: Boolean;
  end;

  TControleAcesso = class
  public
    class function ObterPermissao(AConnection: TFDConnection; AIdPerfil: Integer; const ANomeModulo: string): TPermissao;
    class function TemAcesso(AConnection: TFDConnection; AIdPerfil: Integer; const ANomeModulo: string): Boolean;
  end;

implementation

class function TControleAcesso.ObterPermissao(AConnection: TFDConnection; AIdPerfil: Integer; const ANomeModulo: string): TPermissao;
var
  Qry: TFDQuery;
begin
  Result.CanAccess := False;
  Result.CanInsert := False;
  Result.CanEdit   := False;
  Result.CanDelete := False;

  if (AConnection = nil) then Exit;

  // Se o perfil não foi informado (zerado), assume o perfil 1 para testes
  if AIdPerfil <= 0 then
    AIdPerfil := 1;

  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := AConnection;
    Qry.SQL.Text := 'SELECT P.CAN_ACCESS' + sLineBreak +
                    '     , P.CAN_INSERT' + sLineBreak +
                    '     , P.CAN_EDIT' + sLineBreak +
                    '     , P.CAN_DELETE' + sLineBreak +
                    '  FROM PERMISSAO_PERFIL P' + sLineBreak +
                    '  INNER JOIN MODULO M ON (M.ID_MODULO = P.ID_MODULO)' + sLineBreak +
                    ' WHERE P.ID_PERFIL = :ID_PERFIL' + sLineBreak +
                    '   AND UPPER(TRIM(M.NOME_MODULO)) = UPPER(TRIM(:NOME_MODULO))';
    Qry.ParamByName('ID_PERFIL').AsInteger := AIdPerfil;
    Qry.ParamByName('NOME_MODULO').AsString := ANomeModulo;
    Qry.Open;

    if not Qry.IsEmpty then
    begin
      Result.CanAccess := Qry.FieldByName('CAN_ACCESS').AsString = 'S';
      Result.CanInsert := Qry.FieldByName('CAN_INSERT').AsString = 'S';
      Result.CanEdit   := Qry.FieldByName('CAN_EDIT').AsString = 'S';
      Result.CanDelete := Qry.FieldByName('CAN_DELETE').AsString = 'S';
    end;
  finally
    Qry.Free;
  end;
end;

class function TControleAcesso.TemAcesso(AConnection: TFDConnection; AIdPerfil: Integer; const ANomeModulo: string): Boolean;
var
  Perm: TPermissao;
begin
  Perm := ObterPermissao(AConnection, AIdPerfil, ANomeModulo);
  Result := Perm.CanAccess;
end;

end.
