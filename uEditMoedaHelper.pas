unit uEditMoedaHelper;

interface

uses
  System.SysUtils, Vcl.StdCtrls;

procedure EditMoedaKeyPress(Edit: TEdit; var Key: Char);
procedure EditMoedaExit(Edit: TEdit);

implementation

procedure EditMoedaExit(Edit: TEdit);
begin
  // Ao sair do campo, se for zero, limpa
  if Trim(Edit.Text) = '' then
    Exit;

  if StrToCurrDef(Edit.Text, 0) = 0 then
    Edit.Clear;
end;

procedure EditMoedaKeyPress(Edit: TEdit; var Key: Char);
var
  Digitos: string;
  Valor: Currency;
begin
  // Backspace
  if Key = #8 then
  begin
    Digitos := StringReplace(Edit.Text, FormatSettings.DecimalSeparator, '', [rfReplaceAll]);
    Digitos := StringReplace(Digitos, FormatSettings.ThousandSeparator, '', [rfReplaceAll]);

    if Length(Digitos) > 0 then
      Delete(Digitos, Length(Digitos), 1);

    if Digitos = '' then
      Digitos := '0';

    Valor := StrToCurr(Digitos) / 100;
    Edit.Text := FormatCurr('0.00', Valor);
    Edit.SelStart := Length(Edit.Text);
    Key := #0;
    Exit;
  end;

  // Somente números
  if not (Key in ['0'..'9']) then
  begin
    Key := #0;
    Exit;
  end;

  // Mantém apenas números
  Digitos := StringReplace(Edit.Text, FormatSettings.DecimalSeparator, '', [rfReplaceAll]);
  Digitos := StringReplace(Digitos, FormatSettings.ThousandSeparator, '', [rfReplaceAll]);

  Digitos := Digitos + Key;

  Valor := StrToCurr(Digitos) / 100;

  Edit.Text := FormatCurr('0.00', Valor);
  Edit.SelStart := Length(Edit.Text);

  Key := #0;
end;

end.

