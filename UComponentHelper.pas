unit UComponentHelper;

interface

uses
  System.Classes,
  System.SysUtils,
  Vcl.ExtCtrls,
  Vcl.StdCtrls,
  Vcl.Buttons,
  Vcl.Controls,
  Vcl.Forms,
  Winapi.Windows,
  Winapi.Messages; // Unit necessária para WM_SETREDRAW

type
  TViewUtils = class
  public
    { Método genérico com animação e tratamento de flicker (tremor) }
    class procedure AnimarPainel(APanel: TPanel; AButton: TControl;
      var FWidthStorage: Integer; AMinWidth: Integer = 80);
  end;

implementation

class procedure TViewUtils.AnimarPainel(APanel: TPanel; AButton: TControl;
  var FWidthStorage: Integer; AMinWidth: Integer = 80);
var
  LTargetWidth: Integer;
  LStep: Integer;
const
  ANIMATION_STEP = 20;
  SLEEP_TIME = 1;
begin
  if not Assigned(APanel) then Exit;

  // Habilita o DoubleBuffered para suavizar o desenho interno
  APanel.DoubleBuffered := True;
  if (APanel.Parent <> nil) and (APanel.Parent is TWinControl) then
    TWinControl(APanel.Parent).DoubleBuffered := True;

  // Define o alvo (Abrir ou Fechar)
  if APanel.Width <> AMinWidth then
  begin
    // Fechar: Salva a largura atual antes
    if APanel.Width > AMinWidth then
      FWidthStorage := APanel.Width;
    LTargetWidth := AMinWidth;
    LStep := -ANIMATION_STEP;
  end
  else
  begin
    // Abrir: Recupera a largura salva
    LTargetWidth := FWidthStorage;
    if LTargetWidth <= AMinWidth then LTargetWidth := 250; // Valor padrão caso esteja zerada
    LStep := ANIMATION_STEP;
  end;

  // Bloqueia o redesenho automático do Windows para evitar o tremor
  SendMessage(APanel.Handle, WM_SETREDRAW, Ord(False), 0);
  try
    while Abs(APanel.Width - LTargetWidth) > Abs(LStep) do
    begin
      APanel.Width := APanel.Width + LStep;

      // Força um redesenho rápido e controlado
      SendMessage(APanel.Handle, WM_SETREDRAW, Ord(True), 0);
      APanel.Refresh;
      SendMessage(APanel.Handle, WM_SETREDRAW, Ord(False), 0);

      Sleep(SLEEP_TIME);
      Application.ProcessMessages;
    end;

    // Garante que o painel termine exatamente na largura alvo
    APanel.Width := LTargetWidth;
  finally
    // Desbloqueia definitivamente o redesenho
    SendMessage(APanel.Handle, WM_SETREDRAW, Ord(True), 0);
    APanel.Repaint;
  end;

  // Atualiza o texto do botão conforme o estado final
  if Assigned(AButton) then
  begin
    if APanel.Width <= AMinWidth then
    begin
       if AButton is TButton then TButton(AButton).Caption := 'Exibir'
       else if AButton is TSpeedButton then TSpeedButton(AButton).Caption := 'Exibir';
    end
    else
    begin
       if AButton is TButton then TButton(AButton).Caption := 'Ocultar'
       else if AButton is TSpeedButton then TSpeedButton(AButton).Caption := 'Ocultar';
    end;
  end;
end;

end.
