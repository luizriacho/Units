//unit uAutomacaoSIC;
//
//interface
//
//uses
//  System.SysUtils, System.Classes, Winapi.Windows, Winapi.ShellAPI,
//  System.NetEncoding, Data.DB;
//
//type
//  TAutomacaoSIC = class
//  private
//    class function Pad(const S: string; Len: Integer): string;
//  public
//    class procedure EnviarRelatorioOperador(
//      ANumero: string;
//      ANomeOperador: string;
//      ADesempenho: string;
//      AMediaPontuacao: Double;
//      ATotalOcorrencias : Integer;
//      ARanking: Integer;
//      APeriodoStr: string;
//      ADiasDirecao: Integer;
//      ADataSetDetalhe: TDataSet;
//      ACampoPontuacao: string = 'PONTUACAO_MOVIMENTO'
//    );
//  end;
//
//implementation
//
//{ TAutomacaoSIC }
//
//class function TAutomacaoSIC.Pad(const S: string; Len: Integer): string;
//begin
//  // Garante que o preenchimento de espaços funcione para fontes monoespaçadas
//  Result := S + StringOfChar(' ', Len - Length(S));
//end;
//
//class procedure TAutomacaoSIC.EnviarRelatorioOperador(
//  ANumero: string; ANomeOperador: string; ADesempenho: string;
//  AMediaPontuacao: Double;ATotalOcorrencias, ARanking: Integer; APeriodoStr: string;
//  ADiasDirecao: Integer; ADataSetDetalhe: TDataSet;
//  ACampoPontuacao: string = 'PONTUACAO_MOVIMENTO'
//);
//var
//  LMsg, LUrl, LNumeroLimpo, LParametros, LEmojiSelo: string;
//  i: Integer;
//begin
//  // 1. Limpeza do Telefone
//  LNumeroLimpo := '';
//  for i := 1 to Length(ANumero) do
//    if ANumero[i] in ['0'..'9'] then LNumeroLimpo := LNumeroLimpo + ANumero[i];
//  if (Length(LNumeroLimpo) = 11) then LNumeroLimpo := '55' + LNumeroLimpo;
//
//  // 2. Montagem do Cabeçalho Mestre
//  LMsg := '📊 *RELATÓRIO DE OPERAÇÃO - SIC*' + sLineBreak +
//          '--------------------------------------------' + sLineBreak +
//          '👤 *Operador:* ' + ANomeOperador + sLineBreak +
//          '🏆 *Ranking:* ' + IntToStr(ARanking) + 'º Lugar' + sLineBreak +
//          '📅 *Período:* ' + UpperCase(APeriodoStr) + sLineBreak + sLineBreak +
//          '⭐ *Desempenho:* ' + ADesempenho + sLineBreak +
//          '📈 *Média Pontuação:* ' + FormatFloat('0.00', AMediaPontuacao) + sLineBreak +
//          '☸️ *Dias na Direção:* ' + IntToStr(ADiasDirecao) + ' dias' + sLineBreak +
//          '⚠️ *Total de Ocorrências:* ' + IntToStr(ATotalOcorrencias) + sLineBreak +
//          '--------------------------------------------' + sLineBreak +
//          '📝 *DETALHE DIÁRIO:*' + sLineBreak +
//          // AJUSTE DE ALINHAMENTO: 3 espaços antes da crase para alinhar com o emoji
//          '    `   Dia   |    Veic    |    Pts`' + sLineBreak;
//
//  // 3. Detalhamento
//  ADataSetDetalhe.First;
//  while not ADataSetDetalhe.Eof do
//  begin
//    LEmojiSelo := '⚪';
//    if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'VERDE' then LEmojiSelo := '🟢'
//    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'VERMELHO' then LEmojiSelo := '🔴'
//    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'AMARELO' then LEmojiSelo := '🟡'
//    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'DOURADO' then LEmojiSelo := '⭐';
//
//    // Montagem da linha: Emoji + 1 espaço + crase + dados
//    // O Pad(..., 4) garante que veículos menores não desalinhem a coluna de pontos
//    LMsg := LMsg + LEmojiSelo + ' `' +
//            FormatDateTime('dd', ADataSetDetalhe.FieldByName('DATA').AsDateTime) + '  | ' +
//            Pad(ADataSetDetalhe.FieldByName('ID_VEICULO').AsString, 4) + ' | ' +
//            FormatFloat('0.00', ADataSetDetalhe.FieldByName(ACampoPontuacao).AsFloat) + '`' + sLineBreak;
//
//    ADataSetDetalhe.Next;
//  end;
//
//  LMsg := LMsg + sLineBreak + '_Gerado pelo Sistema SIC_';
//
//  // 4. Execução via Chrome Proxy
//  LUrl := 'https://web.whatsapp.com/send?phone=' + LNumeroLimpo + '&text=' + TNetEncoding.URL.Encode(LMsg);
//  LParametros := '--profile-directory="Default" --ignore-profile-directory-if-not-exists --app=' + LUrl;
//
//  ShellExecute(0, 'open', PChar('C:\Program Files\Google\Chrome\Application\chrome_proxy.exe'),
//               PChar(LParametros), nil, SW_SHOWNORMAL);
//
//  // 5. Automação do Teclado (15s para garantir que o texto longo carregue)
//  Sleep(15000);
//  keybd_event(VK_RETURN, 0, 0, 0);
//  keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
//end;
//
//end.
unit uAutomacaoSIC;

interface

uses
  System.SysUtils, System.Classes, Winapi.Windows, Winapi.ShellAPI,
  System.NetEncoding, Data.DB, Vcl.Forms;

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
      ATotalOcorrencias: Integer;
      ARanking: Integer;
      APeriodoStr: string;
      ADiasDirecao: Integer;
      ADataSetDetalhe: TDataSet;
      ACampoPontuacao: string = 'PONTUACAO_MOVIMENTO'
    );

    class procedure EnviarEscalaOperador(
      ANumero: string;
      ANomeOperador: string;
      ADataEscala: TDate;
      ADataSetEscala: TDataSet
    );

    class procedure EnviarEscalasEmLote(
      ADataEscala: TDate;
      ADataSetFuncionarios: TDataSet;
      ADataSetItinerario: TDataSet
    );
  end;

implementation

{ TAutomacaoSIC }

class function TAutomacaoSIC.Pad(const S: string; Len: Integer): string;
begin
  Result := S + StringOfChar(' ', Len - Length(S));
end;

class procedure TAutomacaoSIC.EnviarRelatorioOperador(
  ANumero: string; ANomeOperador: string; ADesempenho: string;
  AMediaPontuacao: Double; ATotalOcorrencias, ARanking: Integer; APeriodoStr: string;
  ADiasDirecao: Integer; ADataSetDetalhe: TDataSet;
  ACampoPontuacao: string = 'PONTUACAO_MOVIMENTO'
);
var
  LMsg, LUrl, LNumeroLimpo, LParametros, LEmojiSelo: string;
  i: Integer;
begin
  LNumeroLimpo := '';
  for i := 1 to Length(ANumero) do
    if ANumero[i] in ['0'..'9'] then LNumeroLimpo := LNumeroLimpo + ANumero[i];
  if (Length(LNumeroLimpo) = 11) then LNumeroLimpo := '55' + LNumeroLimpo;

  LMsg := '📊 *RELATÓRIO DE OPERAÇÃO - SIC*' + sLineBreak +
          '--------------------------------------------' + sLineBreak +
          '👤 *Operador:* ' + ANomeOperador + sLineBreak +
          '🏆 *Ranking:* ' + IntToStr(ARanking) + 'º Lugar' + sLineBreak +
          '📅 *Período:* ' + UpperCase(APeriodoStr) + sLineBreak + sLineBreak +
          '⭐ *Desempenho:* ' + ADesempenho + sLineBreak +
          '📈 *Média Pontuação:* ' + FormatFloat('0.00', AMediaPontuacao) + sLineBreak +
          '☸️ *Dias na Direção:* ' + IntToStr(ADiasDirecao) + ' dias' + sLineBreak +
          '⚠️ *Total de Ocorrências:* ' + IntToStr(ATotalOcorrencias) + sLineBreak +
          '--------------------------------------------' + sLineBreak +
          '📝 *DETALHE DIÁRIO:*' + sLineBreak +
          '   `   Dia   |    Veic    |    Pts`' + sLineBreak;

  ADataSetDetalhe.First;
  while not ADataSetDetalhe.Eof do
  begin
    LEmojiSelo := '⚪';
    if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'VERDE' then LEmojiSelo := '🟢'
    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'VERMELHO' then LEmojiSelo := '🔴'
    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'AMARELO' then LEmojiSelo := '🟡'
    else if ADataSetDetalhe.FieldByName('COR_SELO').AsString = 'DOURADO' then LEmojiSelo := '⭐';

    LMsg := LMsg + LEmojiSelo + ' `' +
            FormatDateTime('dd', ADataSetDetalhe.FieldByName('DATA').AsDateTime) + '  | ' +
            Pad(ADataSetDetalhe.FieldByName('ID_VEICULO').AsString, 4) + ' | ' +
            FormatFloat('0.00', ADataSetDetalhe.FieldByName(ACampoPontuacao).AsFloat) + '`' + sLineBreak;

    ADataSetDetalhe.Next;
  end;

  LMsg := LMsg + sLineBreak + '_Gerado pelo Sistema SIC_';

  LUrl := 'https://web.whatsapp.com/send?phone=' + LNumeroLimpo + '&text=' + TNetEncoding.URL.Encode(LMsg);
  LParametros := '--profile-directory="Default" --ignore-profile-directory-if-not-exists --app=' + LUrl;

  ShellExecute(0, 'open', PChar('C:\Program Files\Google\Chrome\Application\chrome_proxy.exe'),
                 PChar(LParametros), nil, SW_SHOWNORMAL);

  Sleep(15000);
  keybd_event(VK_RETURN, 0, 0, 0);
  keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
end;

class procedure TAutomacaoSIC.EnviarEscalaOperador(
  ANumero: string;
  ANomeOperador: string;
  ADataEscala: TDate;
  ADataSetEscala: TDataSet
);
var
  LMsg, LUrl, LNumeroLimpo, LParametros: string;
  LLinha, LInicio, LFim, LSaida, LRetorno, LPartida: string;
  i: Integer;
  vBM: TBookmark;
begin
  if (ANumero = '') or (ADataSetEscala = nil) or (ADataSetEscala.IsEmpty) then
    Exit;

  // Limpeza do número
  LNumeroLimpo := '';
  for i := 1 to Length(ANumero) do
    if ANumero[i] in ['0'..'9'] then LNumeroLimpo := LNumeroLimpo + ANumero[i];
  if (Length(LNumeroLimpo) = 11) then LNumeroLimpo := '55' + LNumeroLimpo;

  // Cabeçalho alinhado ao padrão visual do App
  LMsg := '📅 *ESCALA PROGRAMADA*' + sLineBreak +
          '👤 *' + UpperCase(ANomeOperador) + '*' + sLineBreak +
          '📆 *' + FormatDateTime('dd mmm yyyy', ADataEscala) + ' · ' + FormatDateTime('dddd', ADataEscala) + '*' + sLineBreak +
          '--------------------------------------------' + sLineBreak;

  vBM := ADataSetEscala.GetBookmark;
  try
    ADataSetEscala.First;
    while not ADataSetEscala.Eof do
    begin
      LLinha   := Trim(ADataSetEscala.FieldByName('LINHA').AsString);
      LInicio  := Trim(ADataSetEscala.FieldByName('HORA_INICIO').AsString);
      if LInicio = '' then LInicio := '-';

      LFim     := Trim(ADataSetEscala.FieldByName('HORA_FIM').AsString);
      if LFim = '' then LFim := '-';

      LSaida   := Trim(ADataSetEscala.FieldByName('INICIO_VIAGEM').AsString);
      if LSaida = '' then LSaida := '-';

      LRetorno := Trim(ADataSetEscala.FieldByName('PONTO_RETORNO').AsString);
      if LRetorno = '' then LRetorno := '-';

      LPartida := Trim(ADataSetEscala.FieldByName('PONTO_PARTIDA').AsString);
      if LPartida = '' then LPartida := '-';

      // Estrutura espelhada do Card do App
      LMsg := LMsg + '🚌 *LINHA: ' + LLinha + '*' + sLineBreak +
              '⏱️ Início: `' + LInicio + '` | Fim: `' + LFim + '` | Saída: `' + LSaida + '`' + sLineBreak +
              '📍 Partida: ' + LPartida + sLineBreak +
              '🔄 Retorno: ' + LRetorno + sLineBreak +
              '--------------------------------------------' + sLineBreak;

      ADataSetEscala.Next;
    end;
  finally
    if ADataSetEscala.BookmarkValid(vBM) then
      ADataSetEscala.GotoBookmark(vBM);
    ADataSetEscala.FreeBookmark(vBM);
  end;

  LMsg := LMsg + '_Gerado pelo Gerenciador de Escalas - SIC_';

  LUrl := 'https://web.whatsapp.com/send?phone=' + LNumeroLimpo + '&text=' + TNetEncoding.URL.Encode(LMsg);
  LParametros := '--profile-directory="Default" --ignore-profile-directory-if-not-exists --app=' + LUrl;

  ShellExecute(0, 'open', PChar('C:\Program Files\Google\Chrome\Application\chrome_proxy.exe'),
                 PChar(LParametros), nil, SW_SHOWNORMAL);

  Sleep(15000);
  keybd_event(VK_RETURN, 0, 0, 0);
  keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
end;

class procedure TAutomacaoSIC.EnviarEscalasEmLote(
  ADataEscala: TDate;
  ADataSetFuncionarios: TDataSet;
  ADataSetItinerario: TDataSet
);
var
  LChaveFun: Integer;
  LNomeOperador, LWhatsApp: string;
begin
  if (ADataSetFuncionarios = nil) or (ADataSetFuncionarios.IsEmpty) then
    Exit;

  ADataSetFuncionarios.First;
  while not ADataSetFuncionarios.Eof do
  begin
    LChaveFun     := ADataSetFuncionarios.FieldByName('CHAVE_FUN').AsInteger;
    LNomeOperador := ADataSetFuncionarios.FieldByName('NOME').AsString;
    LWhatsApp     := Trim(ADataSetFuncionarios.FieldByName('WHATSAPP').AsString);

    if LWhatsApp <> '' then
    begin
      ADataSetItinerario.Filter   := 'CHAVE_FUN = ' + IntToStr(LChaveFun);
      ADataSetItinerario.Filtered := True;

      if not ADataSetItinerario.IsEmpty then
      begin
        EnviarEscalaOperador(
          LWhatsApp,
          LNomeOperador,
          ADataEscala,
          ADataSetItinerario
        );

        Sleep(3000);
      end;
    end;

    ADataSetFuncionarios.Next;
  end;

  ADataSetItinerario.Filtered := False;
end;

end.
