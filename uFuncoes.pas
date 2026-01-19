unit uFuncoes;

interface

uses Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, Buttons, FMTBcd, SqlExpr,
  Mask, ExtCtrls, Gauges, ComObj, ActiveX, Excel2000, DBCtrls, DateUtils,
  System.IniFiles, Vcl.WinXCtrls, System.StrUtils;

type
  TNumeroStr = string;

const
  Unidades: array [1 .. 19] of TNumeroStr = ('um', 'dois', 'trÍs', 'quatro',
    'cinco', 'seis', 'sete', 'oito', 'nove', 'dez', 'onze', 'doze', 'treze',
    'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove');
  Dezenas: array [1 .. 9] of TNumeroStr = ('dez', 'vinte', 'trinta', 'quarenta',
    'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa');
  Centenas: array [1 .. 9] of TNumeroStr = ('cem', 'duzentos', 'trezentos',
    'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos',
    'novecentos');
  ErrorString = 'Valor fora da faixa';
  Min = 0.01;
  Max = 4294967295.99;
  Moeda = ' real ';
  Moedas = ' reais ';
  Centesimo = ' centavo ';
  Centesimos = ' centavos ';

function ultimoDiaAno(DataAnt: TDate): TDate;
function addAno(DataAnt: TDate): TDate;
function decAno(DataAnt: TDate): TDate;
function UltimoDiaUtilMes(Data: TDateTime; lSabDom: Boolean): TDateTime;
Function PrimeiroDiaMes(Data: TDateTime; lSabDom: Boolean): TDateTime;
function UltimoDiaMesCorrente(Data: TDateTime): TDateTime;
Function PrimeiroDiaMesCorrente(Data: TDateTime): TDateTime;
function DiasPorMes(Ayear, AMonth: Integer): Integer;
function AnoBiSexto(Ayear: Integer): Boolean;
function verMes(Data: TDate): string;
function verMesAbreviado(Data: TDate): string;
function Mes_str(Mes_int: Integer): string;
function ParteTexto(Frase: string): string;
function DifDias(DataVenc: TDateTime; DataAtual: TDateTime): string;
function tirarHoraData(Data: TDate): TDate;
function NumeroParaExtenso(parmNumero: Real): string;
function ConversaoRecursiva(N: LongWord): string;
function CopyLeft(AString: string; ALength: Integer): string;
function CopyRight(AString: string; ALength: Integer): string;
function RemoveAcento(Str: string): string;
function montarData(mesStr: string): TDateTime;
Function CurrentYear: Word;
Function CurrentMonth: Word;
function UltimoDoMes(mesStr: string): TDateTime;
function novaData(DataAnt: TDate): TDate;
function Consistencia(Formulario: Tform): Boolean;

function GetStrNumber(const S: string): string;
function addMes(DataAnt: TDate): TDate;
function decMes(DataAnt: TDate): TDate;
function ValidarEMail(aStr: string): Boolean;
function montarDataCompleta(mesStr: string; anoInt: Integer): TDateTime;
function Mes_Completo(Mes_int: Integer): string;
function parametroData(dia, mes, ano: Word): TDateTime;
function GetTimeInDateTime(Data: TDateTime): Word;
function ColocaTextoEsq(Texto: string; Qtd: Integer; Ch: Char): string;
function ColocaTextoDir(Texto: string; Qtd: Integer; Ch: Char): string;
procedure conectar(empresa: string);
//procedure conferirNovatos;
function tipoDia(dData: TDateTime): Integer;
function textoDireita(sTexto: string; qtde_caract: Integer): string;
function textoEsquerda(sTexto: string; qtde_caract: Integer): string;
function usuarioLogado: String;
function diretorio: string;
function retorna_cpf_operador(chave_f: Integer): string;
function validarData(dImportacao, dParametro: TDate): Boolean;
function PreencherComChar(const Texto: string; TamanhoTotal: Integer;
  const Caractere: Char): string;

//
// Retorna uma parte de um texto antes de um caractere especificado
//

implementation

uses uPrincipal, udmDados;
function PreencherComChar(const Texto: string; TamanhoTotal: Integer;
  const Caractere: Char): string;
begin
  if length(Texto) < TamanhoTotal then
    result := Texto + StringOfChar(Caractere, TamanhoTotal - length(Texto))
  else
    result := copy(Texto, 1, TamanhoTotal);
end;


function retorna_cpf_operador(chave_f: Integer): string;
begin
  Result := dmDados.con.ExecSQLScalar
    ('select substring(f.cpf FROM 1 FOR 6) from funcionario where chave_fun=:chave_fun ',
    [chave_f]);
end;


function validarData(dImportacao, dParametro: TDate): Boolean;
begin
  if dImportacao >= dParametro then
    Result := True
  else
    Result := False;
end;


function diretorio: string;
begin
  if DirectoryExists
    ('G:\Outros computadores\Meu computador (1)\Arquivos Importar\') then
    result := ('G:\Outros computadores\Meu computador (1)\Arquivos Importar\')
  else
    result := ('C:\Arquivos Importar\');
end;

function usuarioLogado: String;
var
  I: DWord;
  user: string;
begin
  I := 255;
  SetLength(user, I);
  Windows.GetUserName(PChar(user), I);
  user := string(PChar(user));
  result := user;
end;
function textoDireita(sTexto: string; qtde_caract: Integer): string;
var
  AuxStr: String;
begin
  AuxStr := sTexto;
  result := RightStr(AuxStr, qtde_caract); // Ir· Copiar a palavra Delphi
end;

function textoEsquerda(sTexto: string; qtde_caract: Integer): string;
var
  AuxStr: String;
begin
  AuxStr := sTexto;
  result := LeftStr(AuxStr, qtde_caract); // Ir· Copiar a palavra Delphi
end;

function tipoDia(dData: TDateTime): Integer;
{ Verifica se uma data informada cai em dia util, sabado ou domingo }
begin
  if DayOfWeek(dData) = 7 then
    result := 2
  else if DayOfWeek(dData) = 1 then
    result := 3
  else
    result := 1;
end;

function GetTimeInDateTime(Data: TDateTime): Word;
begin
  Result := (HourOf(Data));
end;

Function PrimeiroDiaMesCorrente(Data: TDateTime): TDateTime;
var
  ano, mes, dia: Word;
begin
  DecodeDate(Data, ano, mes, dia);
  dia := 1;
  PrimeiroDiaMesCorrente := EncodeDate(ano, mes, dia);
end;

Function UltimoDiaMesCorrente(Data: TDateTime): TDateTime;
var
  ano, mes, dia: Word;
begin
  DecodeDate(Data, ano, mes, dia);
  dia := DiasPorMes(ano, mes);
  UltimoDiaMesCorrente := EncodeDate(ano, mes, dia);
end;

Function PrimeiroDiaMes(Data: TDateTime; lSabDom: Boolean): TDateTime;
var
  ano, mes, dia: Word;
  DiaDaSemana: Integer;
begin
  DecodeDate(Data, ano, mes, dia);
  dia := 1;
  if lSabDom Then
  begin
    DiaDaSemana := DayOfWeek(Data);
    if DiaDaSemana = 1 Then
      dia := 2
    else if DiaDaSemana = 7 Then
      dia := 3;
  end;
  PrimeiroDiaMes := EncodeDate(ano, mes, dia);
end;

function UltimoDiaUtilMes(Data: TDateTime; lSabDom: Boolean): TDateTime;
var
  ano, mes, dia: Word;
  AuxData: TDateTime;
  DiaDaSemana: Integer;
begin
  AuxData := Data;
  if lSabDom then
  begin
    DecodeDate(AuxData, ano, mes, dia);
    DiaDaSemana := DayOfWeek(AuxData);
    if DiaDaSemana = 1 then
      dia := dia - 2
    else if DiaDaSemana = 7 then
      Dec(dia);
    AuxData := EncodeDate(ano, mes, dia);
  end;
  UltimoDiaUtilMes := AuxData;
end;

function AnoBiSexto(Ayear: Integer): Boolean;
begin
  // Verifica se o ano È Bi-Sexto
  Result := (Ayear mod 4 = 0) and ((Ayear mod 100 <> 0) or (Ayear mod 400 = 0));
end;

function DiasPorMes(Ayear, AMonth: Integer): Integer;
const
  DaysInMonth: array [1 .. 12] of Integer = (31, 28, 31, 30, 31, 30, 31, 31, 30,
    31, 30, 31);
begin
  Result := DaysInMonth[AMonth];
  if (AMonth = 2) and AnoBiSexto(Ayear) then
    Inc(Result);
end;

function verMes(Data: TDate): string;
var
  ano, mes, dia: Word;
  mensal: array [1 .. 12] of string;
begin
  DecodeDate(Data, ano, mes, dia);
  mensal[1] := 'Janeiro';
  mensal[2] := 'Fevereiro';
  mensal[3] := 'MarÁo';
  mensal[4] := 'Abril';
  mensal[5] := 'Maio';
  mensal[6] := 'Junho';
  mensal[7] := 'Julho';
  mensal[8] := 'Agosto';
  mensal[9] := 'Setembro';
  mensal[10] := 'Outubro';
  mensal[11] := 'Novembro';
  mensal[12] := 'Dezembro';
  Result := mensal[mes] + '/' + IntToStr(ano);
end;

function verMesAbreviado(Data: TDate): string;
var
  ano, mes, dia: Word;
  mensal: array [1 .. 12] of string;
begin
  DecodeDate(Data, ano, mes, dia);
  mensal[1] := 'Jan';
  mensal[2] := 'Fev';
  mensal[3] := 'Mar';
  mensal[4] := 'Abr';
  mensal[5] := 'Mai';
  mensal[6] := 'Jun';
  mensal[7] := 'Jul';
  mensal[8] := 'Ago';
  mensal[9] := 'Set';
  mensal[10] := 'Out';
  mensal[11] := 'Nov';
  mensal[12] := 'Dez';
  Result := mensal[mes] + '/' + IntToStr(ano);
end;

function Mes_str(Mes_int: Integer): string;
var
  mensal: array [1 .. 12] of string;
begin

  mensal[1] := 'Jan';
  mensal[2] := 'Fev';
  mensal[3] := 'Mar';
  mensal[4] := 'Abr';
  mensal[5] := 'Mai';
  mensal[6] := 'Jun';
  mensal[7] := 'Jul';
  mensal[8] := 'Ago';
  mensal[9] := 'Set';
  mensal[10] := 'Out';
  mensal[11] := 'Nov';
  mensal[12] := 'Dez';
  Result := mensal[Mes_int];
end;

function Mes_Completo(Mes_int: Integer): string;
var
  mensal: array [1 .. 12] of string;
begin

  mensal[1] := 'JANEIRO';
  mensal[2] := 'FEVEREIRO';
  mensal[3] := 'MAR«O';
  mensal[4] := 'ABRIL';
  mensal[5] := 'MAIO';
  mensal[6] := 'JUNHO';
  mensal[7] := 'JULHO';
  mensal[8] := 'AGOSTO';
  mensal[9] := 'SETEMBRO';
  mensal[10] := 'OUTUBRO';
  mensal[11] := 'NOVEMBRO';
  mensal[12] := 'DEZEMBRO';
  Result := mensal[Mes_int];
end;

function ParteTexto(Frase: string): string;
//
// Retorna uma parte de um texto antes de um caractere especificado
//
var
  i, Max: Integer;
  buff: string;
begin
  i := 1;
  buff := '';
  Max := length(' ');
  while (i <= length(Frase)) and (buff <> ' ') do
  begin
    buff := buff + Frase[i];
    if i > 12 then
    begin
      if length(buff) > Max then
      begin
        buff := copy(buff, 2, Max);
      end;
    end;
    Inc(i);
  end;
  if buff = ' ' then
  begin
    Result := copy(Frase, 1, i - Max - 1);
    Frase := copy(Frase, i, length(Frase) + 1 - i);
  end
  else
  begin
    Result := Frase;
    Frase := '';
  end;
end;

function DifDias(DataVenc: TDateTime; DataAtual: TDateTime): string;
{ Retorna a diferenca de dias entre duas datas }
var
  Data: TDateTime;
  dia, mes, ano: Word;
begin
  if DataAtual < DataVenc then
  begin
    Result := 'A data data atual n„o pode ser menor que a data inicial';
  end
  else
  begin
    Data := DataAtual - DataVenc;
    DecodeDate(Data, ano, mes, dia);
    Result := FloatToStr(Data);
  end;
end;

function tirarHoraData(Data: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(Data, ano, mes, dia);
  Result := EncodeDate(ano, mes, dia);
end;

function NumeroParaExtenso(parmNumero: Real): string;
begin
  if (parmNumero >= Min) and (parmNumero <= Max) then
  begin
    { Tratar reais }
    Result := ConversaoRecursiva(Round(Int(parmNumero)));
    if Round(Int(parmNumero)) = 1 then
      Result := Result + Moeda
    else if Round(Int(parmNumero)) <> 0 then
      Result := Result + Moedas;
    { Tratar centavos }
    if not(Frac(parmNumero) = 0.00) then
    begin
      if Round(Int(parmNumero)) <> 0 then
        Result := Result + ' e ';
      Result := Result + ConversaoRecursiva(Round(Frac(parmNumero) * 100));
      if (Round(Frac(parmNumero) * 100) = 1) then
        Result := Result + Centesimo
      else
        Result := Result + Centesimos;
    end;
  end
  else
    raise ERangeError.CreateFmt('%g ' + ErrorString + ' %g..%g',
      [parmNumero, Min, Max]);
end;

function ConversaoRecursiva(N: LongWord): string;
begin
  case N of
    1 .. 19:
      Result := Unidades[N];
    20, 30, 40, 50, 60, 70, 80, 90:
      Result := Dezenas[N div 10] + ' ';
    21 .. 29, 31 .. 39, 41 .. 49, 51 .. 59, 61 .. 69, 71 .. 79,
      81 .. 89, 91 .. 99:
      Result := Dezenas[N div 10] + ' e ' + ConversaoRecursiva(N mod 10);
    100, 200, 300, 400, 500, 600, 700, 800, 900:
      Result := Centenas[N div 100] + ' ';
    101 .. 199:
      Result := ' cento e ' + ConversaoRecursiva(N mod 100);
    201 .. 299, 301 .. 399, 401 .. 499, 501 .. 599, 601 .. 699, 701 .. 799,
      801 .. 899, 901 .. 999:
      Result := Centenas[N div 100] + ' e ' + ConversaoRecursiva(N mod 100);
    1000 .. 999999:
      Result := ConversaoRecursiva(N div 1000) + ' mil ' +
        ConversaoRecursiva(N mod 1000);
    1000000 .. 1999999:
      Result := ConversaoRecursiva(N div 1000000) + ' milh„o ' +
        ConversaoRecursiva(N mod 1000000);
    2000000 .. 999999999:
      Result := ConversaoRecursiva(N div 1000000) + ' milhıes ' +
        ConversaoRecursiva(N mod 1000000);
    1000000000 .. 1999999999:
      Result := ConversaoRecursiva(N div 1000000000) + ' bilh„o ' +
        ConversaoRecursiva(N mod 1000000000);
    2000000000 .. 4294967295:
      Result := ConversaoRecursiva(N div 1000000000) + ' bilhıes ' +
        ConversaoRecursiva(N mod 1000000000);
  end;
end;
{ CopyLeft }

function CopyLeft(AString: string; ALength: Integer): string;
var
  i: Integer;
begin
  { Retorna uma SubString da direita para esquerda no tamamnho desejado }
  // substitui pq. no delhpi 5 n„o tem StrUtils
  for i := 1 to length(AString) do
    if i > ALength then
      Break
    else
      Result := Result + AString[i];
end;

{ CopyRight }

function CopyRight(AString: string; ALength: Integer): string;
var
  i: Integer;
begin
  { Retorna uma SubString da esquerda para direita no tamamnho desejado }
  // substitui pq. no delhpi 5 n„o tem StrUtils
  for i := length(AString) downto 1 do
    if (length(AString) - i) >= ALength then
      Break
    else
      Result := AString[i] + Result;
end;

function RemoveAcento(Str: string): string;
const

  ComAcento = '‡‚ÍÙ˚„ı·ÈÌÛ˙Á¸¿¬ ‘€√’¡…Õ”⁄«‹';

  SemAcento = 'aaeouaoaeioucuAAEOUAOAEIOUCU';

var

  x: Integer;

begin;

  for x := 1 to length(Str) do

    if Pos(Str[x], ComAcento) <> 0 then

      Str[x] := SemAcento[Pos(Str[x], ComAcento)];

  Result := Str;

end;

function montarData(mesStr: string): TDateTime;
var
  wYear: TSystemTime;
  ano, mes, dia: Word;
  inteiro: Integer;
begin
  if mesStr = 'Janeiro' then
    inteiro := 1
  else if mesStr = 'Fevereiro' then
    inteiro := 2
  else if mesStr = 'MarÁo' then
    inteiro := 3
  else if mesStr = 'Abril' then
    inteiro := 4
  else if mesStr = 'Maio' then
    inteiro := 5
  else if mesStr = 'Junho' then
    inteiro := 6
  else if mesStr = 'Julho' then
    inteiro := 7
  else if mesStr = 'Agosto' then
    inteiro := 8
  else if mesStr = 'Setembro' then
    inteiro := 9
  else if mesStr = 'Outubro' then
    inteiro := 10
  else if mesStr = 'Novembro' then
    inteiro := 11
  else if mesStr = 'Dezembro' then
    inteiro := 12;
  Result := EncodeDate(CurrentYear, inteiro, 1);
end;

function UltimoDoMes(mesStr: string): TDateTime;
var
  wYear: TSystemTime;
  ano, mes, dia: Word;
  inteiro: Integer;
begin
  if mesStr = 'Janeiro' then
    inteiro := 1
  else if mesStr = 'Fevereiro' then
    inteiro := 2
  else if mesStr = 'MarÁo' then
    inteiro := 3
  else if mesStr = 'Abril' then
    inteiro := 4
  else if mesStr = 'Maio' then
    inteiro := 5
  else if mesStr = 'Junho' then
    inteiro := 6
  else if mesStr = 'Julho' then
    inteiro := 7
  else if mesStr = 'Agosto' then
    inteiro := 8
  else if mesStr = 'Setembro' then
    inteiro := 9
  else if mesStr = 'Outubro' then
    inteiro := 10
  else if mesStr = 'Novembro' then
    inteiro := 11
  else if mesStr = 'Dezembro' then
    inteiro := 12;
  Result := EncodeDate(CurrentYear, inteiro, DiasPorMes(CurrentYear, inteiro));
end;

Function CurrentYear: Word;
var
  SystemTime: TSystemTime;
begin
  GetLocalTime(SystemTime);
  Result := SystemTime.wYear;
end;

Function CurrentMonth: Word;
var
  SystemTime: TSystemTime;
begin
  GetLocalTime(SystemTime);
  Result := SystemTime.wMonth;
end;

function novaData(DataAnt: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(DataAnt, ano, mes, dia);
  if mes = 12 then
  begin
    ano := ano + 1;
    mes := 1;
    dia := 12
  end
  else
  begin
    ano := ano;
    mes := mes + 1;
    dia := 12
  end;
  Result := EncodeDate(ano, mes, dia);
end;

function addAno(DataAnt: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(DataAnt, ano, mes, dia);
  ano := ano + 1;
  mes := 1;
  dia := 1;
  Result := EncodeDate(ano, mes, dia);
end;

function decAno(DataAnt: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(DataAnt, ano, mes, dia);
  ano := ano - 1;
  mes := 1;
  dia := 1;
  Result := EncodeDate(ano, mes, dia);
end;

function addMes(DataAnt: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(DataAnt, ano, mes, dia);
  if mes = 12 then
  begin
    ano := ano + 1;
    mes := 1;
    dia := 1;
  end
  else
  begin
    ano := ano;
    mes := mes + 1;
    dia := 1
  end;
  Result := EncodeDate(ano, mes, dia);
end;

function decMes(DataAnt: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(DataAnt, ano, mes, dia);
  if mes = 1 then
  begin
    ano := ano - 1;
    mes := 12;
    dia := 1
  end
  else
  begin
    ano := ano;
    mes := mes - 1;
    dia := 1
  end;
  Result := EncodeDate(ano, mes, dia);
end;

function ultimoDiaAno(DataAnt: TDate): TDate;
var
  dia, mes, ano: Word;
begin
  DecodeDate(DataAnt, ano, mes, dia);
  ano := ano;
  mes := 12;
  dia := 31;
  Result := EncodeDate(ano, mes, dia);
end;

function Consistencia(Formulario: Tform): Boolean;
var
  i: Integer;
  Resposta: Boolean;
  Componente: TEdit;
begin
  // Inicializa a resposta
  Resposta := False;
  // Executa uma repetiÁ„o em todos os componentes
  for i := 0 to Formulario.ComponentCount - 1 do
  begin
    // Verifica se o componente È um editBox
    if Formulario.Components[i] is TEdit then
    begin
      // Grava o componente em uma vari·vel
      Componente := Formulario.Components[i] as TEdit;
      // Verifica se o valor est· vazio
      if Componente.Text = '' then
      begin
        Resposta := True;
        Break;
      end; // Fim do if
    end; // Fim do if
  end; // Fim do for
  Result := Resposta;
  case MessageDlg('Deseja se conectar a rede remota?', mtConfirmation,
    [mbYes, mbNo], 0) of
    mrYes:
      begin

      end;
  end;
end;



function GetStrNumber(const S: string): string;
var
  vText: PChar;
begin
  vText := PChar(S);
  Result := '';

  while (vText^ <> #0) do
  begin
{$IFDEF UNICODE}
    if CharInSet(vText^, ['0' .. '9']) then
{$ELSE}
    if vText^ in ['0' .. '9'] then
{$ENDIF}
      Result := Result + vText^;
    Inc(vText);
  end;
end;

function ValidarEMail(aStr: string): Boolean;
begin
  aStr := Trim(UpperCase(aStr));
  if Pos('@', aStr) > 1 then
  begin
    Delete(aStr, 1, Pos('@', aStr));
    Result := (length(aStr) > 0) and (Pos('.', aStr) > 2);
  end
  else
    Result := False;
end;

function montarDataCompleta(mesStr: string; anoInt: Integer): TDateTime;
var
  wYear: TSystemTime;
  ano, mes, dia: Word;
  inteiro: Integer;
begin
  if (mesStr = 'Janeiro') or (mesStr = 'Jan') then
    inteiro := 1
  else if (mesStr = 'Fevereiro') or (mesStr = 'Fev') then
    inteiro := 2
  else if (mesStr = 'MarÁo') or (mesStr = 'Mar') then
    inteiro := 3
  else if (mesStr = 'Abril') or (mesStr = 'Abr') then
    inteiro := 4
  else if (mesStr = 'Maio') or (mesStr = 'Mai') then
    inteiro := 5
  else if (mesStr = 'Junho') or (mesStr = 'Jun') then
    inteiro := 6
  else if (mesStr = 'Julho') or (mesStr = 'Jul') then
    inteiro := 7
  else if (mesStr = 'Agosto') or (mesStr = 'Ago') then
    inteiro := 8
  else if (mesStr = 'Setembro') or (mesStr = 'Set') then
    inteiro := 9
  else if (mesStr = 'Outubro') or (mesStr = 'Out') then
    inteiro := 10
  else if (mesStr = 'Novembro') or (mesStr = 'Nov') then
    inteiro := 11
  else if (mesStr = 'Dezembro') or (mesStr = 'Dez') then
    inteiro := 12;
  Result := EncodeDate(anoInt, inteiro, 1);
end;

function parametroData(dia, mes, ano: Word): TDateTime;
begin
  Result := EncodeDate(ano, mes, dia);
end;

function ColocaTextoEsq(Texto: string; Qtd: Integer; Ch: Char): string;
var
  x: Integer;
begin
  if Ch = '' then
    Ch := Chr(32) { EspaÁo }
    { endif };

  if length(Texto) > Qtd then
    Result := copy(Texto, 0, Qtd)
  else
  begin
    x := length(Texto);
    for Qtd := x to Qtd - 1 do
    begin
      Texto := Texto + Ch;
    end;
    Result := Texto;
  end
  { endif };
end;

function ColocaTextoDir(Texto: string; Qtd: Integer; Ch: Char): string;
var
  x: Integer;
  Str: string;
begin
  if length(Texto) > Qtd then
    Result := copy(Texto, (length(Texto) - Qtd) + 1, length(Texto))
  else
  begin
    Str := '';
    for x := length(Texto) to Qtd - 1 do
    begin
      Str := Str + Ch;
    end;
    Result := Str + Texto;
  end
  { endif };
end;

procedure conectar(empresa: string);
var
  arqIni: TIniFile;
begin
  if empresa = 'BM' then
  BEGIN
    with frmprincipal do
    begin

      with dmDados do
      begin

        lblUser.Caption := 'Us˙ario Ativo: ' +
          (UserControl1.CurrentUser.UserName);

        if (UserControl1.CurrentUser.UserLogin) = 'operador' then
        begin
          actConsultaOperadoresExecute(frmprincipal);
          BorderIcons := BorderIcons - [biSystemMenu];
          spMenu.CloseStyle := svcCollapse;
          spMenu.Opened := False;
        end
        else
        begin
          spMenu.CloseStyle := svcCompact;
          spMenu.Opened := True;
        end;
        if (UserControl1.CurrentUser.UserLogin) = 'abastecimento' then
        begin
          actAbastecimentoManualExecute(frmprincipal);
          BorderIcons := BorderIcons - [biSystemMenu];
        end;

        abertura := 'LOCAL';
        if ((frmprincipal.UserControl1.CurrentUser.UserLogin = 'luiz') or
          (frmprincipal.UserControl1.CurrentUser.UserLogin = 'evandro')) then
        begin
          case MessageDlg('Deseja se conectar a rede remota?', mtConfirmation,
            [mbYes, mbNo], 0) of
            mrYes:
              begin
                abertura := 'ACESSO REMOTO';
                arqIni := TIniFile.Create(ExtractFilePath(ParamStr(0)) +
                  'sic.ini');
                try
                  ADOConnection1.ConnectionString :=
                    arqIni.ReadString('EMPRESA', empresa, '');
                  ADOConnection1.Connected := True;
                except
                  abort;
                end;
                conferirNovatos;
                pnlRodape.Caption :=
                  'SIC - Sistema Integrado de Controle de ManutenÁ„o - Desenvolvido por JosÈ Luiz - (31) 99141-6171 - ACESSO REMOTO - VIA«√O BELO MONTE';

              end;
            mrNo:
              begin
                abertura := 'LOCAL';
              end;
          end;
        end;

        qryParametros.Close;
        qryParametros.Open;
      end;

    end;
  end;

end;

//procedure conferirNovatos;
//begin
//  with dmDados do
//  begin
//    qryMotorista.Close;
//    qryMotorista.Open;
////    qryMotoristaDesligado.Close;
////    qryMotoristaDesligado.Open;
//
//    if qryMotorista.RecordCount > 0 then
//    begin
//      qryMotorista.First;
//      while not qryMotorista.Eof do
//      begin
//        try
//          qryLocalizaChaveFunc.Close;
//          qryLocalizaChaveFunc.Params[0].AsInteger :=
//            qryMotoristaID_FUNCIONARIO.AsInteger;
//          qryLocalizaChaveFunc.Open;
//          if qryLocalizaChaveFunc.RecordCount = 0 then
//          begin
//            qryInserirMotorista.Params[0].AsInteger :=
//              qryMotoristaID_FUNCIONARIO.AsInteger;
//            qryInserirMotorista.Params[1].AsString := qryMotoristaNOME.AsString;
//            qryInserirMotorista.Params[2].AsInteger :=
//              qryMotoristaID_CARGO.AsInteger;
//            qryInserirMotorista.Params[3].AsString := qryMotoristaCPF.AsString;
//            qryInserirMotorista.Params[4].AsDate :=
//              qryMotoristaDATA_ADMISSAO.AsDateTime;
//            qryInserirMotorista.Params[5].AsString :=
//              qryMotoristaTELEFONE.AsString;
//            qryInserirMotorista.ExecSQL;
//            qryInserirMotorista.CommitUpdates;
//          end;
//          // end;
//          qryMotorista.Next;
//        except
//          qryMotorista.Next;
//        end;
//      end;
//    end;
////    if qryMotoristaDesligado.RecordCount > 0 then
////    begin
////      qryMotoristaDesligado.First;
////      while not qryMotoristaDesligado.Eof do
////      begin
////        try
////          con.StartTransaction;
////          con.ExecSQL
////            ('update funcionario set data_demissao=:data_demissao,situacao =:sitInativo where chapa =:chapa and situacao =:sitAtivo and chapa > 1',
////            [qryMotoristaDesligadoDATA_DESLIGAMENTO.AsDateTime, 'N',
////            qryMotoristaDesligadoID_FUNCIONARIO.AsInteger, 'S']);
////          con.Commit;
////          // end;
////          qryMotoristaDesligado.Next;
////        except
////          con.Rollback;
////          qryMotoristaDesligado.Next;
////        end;
////      end;
//    end;
//
////    qryMotorista.Close;
////    qryMotoristaDesligado.Close;
////  end;
//end;

end.
