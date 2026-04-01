unit uAutomacaoSIC;

interface

uses
  System.SysUtils
  , System.Classes
  , Winapi.Windows
  , Winapi.ShellAPI
  , System.NetEncoding;

type
  TAutomacaoSIC = class
  public
    class procedure EnviarRelatorioOperador(
      ANumero: string
      ; ANomeOperador: string
      ; ADesempenho: string
      ; AMediaPontuacao: Double
      ; ARanking: Integer
      ; ADiasTrabalhados: string  // Novo parâmetro
      ; APeriodo: string          // Novo parâmetro
    );
  end;

implementation

{ TAutomacaoSIC }

class procedure TAutomacaoSIC.EnviarRelatorioOperador(
  ANumero: string
  ; ANomeOperador: string
  ; ADesempenho: string
  ; AMediaPontuacao: Double
  ; ARanking: Integer
  ; ADiasTrabalhados: string
  ; APeriodo: string
);
var
  LMsg, LUrl, LNumeroLimpo, LParametros: string;
  i: Integer;
begin
  LNumeroLimpo := '';
  for i := 1 to Length(ANumero) do
    if ANumero[i] in ['0'..'9'] then
      LNumeroLimpo := LNumeroLimpo + ANumero[i];

  if (Length(LNumeroLimpo) = 11) then
    LNumeroLimpo := '55' + LNumeroLimpo;

  LMsg := '📊 *RELATÓRIO DE OPERAÇÃO - SIC*' + sLineBreak +
          '----------------------------------------' + sLineBreak +
          '👤 *Operador:* ' + ANomeOperador + sLineBreak +
          '🏆 *Ranking:* ' + IntToStr(ARanking) + 'º Lugar' + sLineBreak +
          '📅 *Período:* ' + APeriodo + sLineBreak +
          sLineBreak +
          '⭐ *Desempenho:* ' + ADesempenho + sLineBreak +
          '📈 *Média Pontuação:* ' + FormatFloat('0.00', AMediaPontuacao) + sLineBreak +
          '☸️ *Dias na Direção:* ' + ADiasTrabalhados + ' dias' + sLineBreak +
          '----------------------------------------' + sLineBreak +
          ' _Gerado pelo Sistema SIC_';

  LUrl := 'https://web.whatsapp.com/send?phone=' + LNumeroLimpo + '&text=' + TNetEncoding.URL.Encode(LMsg);

  // Usamos o --app para manter o visual de programa e não de navegador
  LParametros := '--app=' + LUrl;

  ShellExecute(0, 'open',
    PChar('C:\Program Files\Google\Chrome\Application\chrome.exe'), // Usando o exe direto
    PChar(LParametros),
    nil, SW_SHOWNORMAL);

  // Espera o carregamento da conversa
  Sleep(5000);

  // Envia o Enter
  keybd_event(VK_RETURN, 0, 0, 0);
  keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);

  // Opcional: Alt+F4 para fechar a janela após enviar, evitando acúmulo
  Sleep(500);
  keybd_event(VK_MENU, 0, 0, 0);
  keybd_event(VK_F4, 0, 0, 0);
  keybd_event(VK_F4, 0, KEYEVENTF_KEYUP, 0);
  keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
end;
end.
