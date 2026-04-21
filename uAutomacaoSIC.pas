unit uAutomacaoSIC;

interface

uses
  System.SysUtils, System.Classes, Winapi.Windows, Winapi.ShellAPI,
  System.NetEncoding, Data.DB;

type
  TAutomacaoSIC = class
  private
    class function Pad(const S: string; Len: Integer): string;
  public
    class procedure EnviarRelatorioOperador(
      ANumero: string;
      ANomeOperador: string;
      ADesempenho: string;
      AMediaPontuacao: Double;
      ARanking: Integer;
      APeriodoStr: string;
      ADiasDirecao: Integer;
      ADataSetDetalhe: TDataSet;
      ACampoPontuacao: string = 'PONTUACAO_MOVIMENTO'
    );
  end;

implementation

{ TAutomacaoSIC }

class function TAutomacaoSIC.Pad(const S: string; Len: Integer): string;
begin
  // Garante que o preenchimento de espaços funcione para fontes monoespaçadas
  Result := S + StringOfChar(' ', Len - Length(S));
end;

class procedure TAutomacaoSIC.EnviarRelatorioOperador(
  ANumero: string; ANomeOperador: string; ADesempenho: string;
  AMediaPontuacao: Double; ARanking: Integer; APeriodoStr: string;
  ADiasDirecao: Integer; ADataSetDetalhe: TDataSet;
  ACampoPontuacao: string = 'PONTUACAO_MOVIMENTO'
);
var
  LMsg, LUrl, LNumeroLimpo, LParametros, LEmojiSelo: string;
  i: Integer;
begin
  // 1. Limpeza do Telefone
  LNumeroLimpo := '';
  for i := 1 to Length(ANumero) do
    if ANumero[i] in ['0'..'9'] then LNumeroLimpo := LNumeroLimpo + ANumero[i];
  if (Length(LNumeroLimpo) = 11) then LNumeroLimpo := '55' + LNumeroLimpo;

  // 2. Montagem do Cabeçalho Mestre
  LMsg := '📊 *RELATÓRIO DE OPERAÇÃO - SIC*' + sLineBreak +
          '--------------------------------------------' + sLineBreak +
          '👤 *Operador:* ' + ANomeOperador + sLineBreak +
          '🏆 *Ranking:* ' + IntToStr(ARanking) + 'º Lugar' + sLineBreak +
          '📅 *Período:* ' + UpperCase(APeriodoStr) + sLineBreak + sLineBreak +
          '⭐ *Desempenho:* ' + ADesempenho + sLineBreak +
          '📈 *Média Pontuação:* ' + FormatFloat('0.00', AMediaPontuacao) + sLineBreak +
          '☸️ *Dias na Direção:* ' + IntToStr(ADiasDirecao) + ' dias' + sLineBreak +
          '--------------------------------------------' + sLineBreak +
          '📝 *DETALHE DIÁRIO:*' + sLineBreak +
          // AJUSTE DE ALINHAMENTO: 3 espaços antes da crase para alinhar com o emoji
          '    `   Dia   |    Veic    |    Pts`' + sLineBreak;

  // 3. Detalhamento
  ADataSetDetalhe.First;
  while not ADataSetDetalhe.Eof do
  begin
    LEmojiSelo := '⚪';
    if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'VERDE' then LEmojiSelo := '🟢'
    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'VERMELHO' then LEmojiSelo := '🔴'
    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'AMARELO' then LEmojiSelo := '🟡'
    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'DOURADO' then LEmojiSelo := '⭐';

    // Montagem da linha: Emoji + 1 espaço + crase + dados
    // O Pad(..., 4) garante que veículos menores não desalinhem a coluna de pontos
    LMsg := LMsg + LEmojiSelo + ' `' +
            FormatDateTime('dd', ADataSetDetalhe.FieldByName('DATA').AsDateTime) + '  | ' +
            Pad(ADataSetDetalhe.FieldByName('ID_VEICULO').AsString, 4) + ' | ' +
            FormatFloat('0.00', ADataSetDetalhe.FieldByName(ACampoPontuacao).AsFloat) + '`' + sLineBreak;

    ADataSetDetalhe.Next;
  end;

  LMsg := LMsg + sLineBreak + '_Gerado pelo Sistema SIC_';

  // 4. Execução via Chrome Proxy
  LUrl := 'https://web.whatsapp.com/send?phone=' + LNumeroLimpo + '&text=' + TNetEncoding.URL.Encode(LMsg);
  LParametros := '--profile-directory="Default" --ignore-profile-directory-if-not-exists --app=' + LUrl;

  ShellExecute(0, 'open', PChar('C:\Program Files\Google\Chrome\Application\chrome_proxy.exe'),
               PChar(LParametros), nil, SW_SHOWNORMAL);

  // 5. Automação do Teclado (15s para garantir que o texto longo carregue)
  Sleep(15000);
  keybd_event(VK_RETURN, 0, 0, 0);
  keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
end;

end.
