unit UOperacaoMestreService;

interface

uses
  System.SysUtils, FireDAC.Comp.Client, Data.DB, System.Variants;

type
  TOperacaoMestreService = class
  public
    class procedure ExecutarConsultaMestre(
      AQry: TFDQuery
      ; const AMotId: Variant
      ; const ALnhId: Variant
      ; const ASitId: Variant
      ; const APosicaoId: Variant
      ; const AFiltrarWhatsapp: Integer
      ; const APeriodo: TDate
      ; const AMpI: Double
      ; const AMpF: Double
      ; const ADiasTrabalhados: Integer
      ; const AConsistir: Integer
      ; const AGarId: Variant
    );
  end;

implementation

class procedure TOperacaoMestreService.ExecutarConsultaMestre(
  AQry: TFDQuery
  ; const AMotId: Variant
  ; const ALnhId: Variant
  ; const ASitId: Variant
  ; const APosicaoId: Variant
  ; const AFiltrarWhatsapp: Integer
  ; const APeriodo: TDate
  ; const AMpI: Double
  ; const AMpF: Double
  ; const ADiasTrabalhados: Integer
  ; const AConsistir: Integer
  ; const AGarId: Variant
);
begin
  AQry.Close;

  // Atribuindo os parâmetros nominalmente
  AQry.ParamByName('mot.id').Value           := AMotId;
  AQry.ParamByName('lnh.id').Value           := ALnhId;
  AQry.ParamByName('sit.id').Value           := ASitId;
  AQry.ParamByName('posicao.id').Value       := APosicaoId;
  AQry.ParamByName('filtrar_whatsapp').Value := AFiltrarWhatsapp;
  AQry.ParamByName('periodo').AsDate         := APeriodo;
  AQry.ParamByName('mpI').Value              := AMpI;
  AQry.ParamByName('mpF').Value              := AMpF;
  AQry.ParamByName('diasTrabalhados').Value  := ADiasTrabalhados;
  AQry.ParamByName('consistir').Value        := AConsistir;
  AQry.ParamByName('gar.id').Value           := AGarId;

  AQry.Open;
end;

end.
