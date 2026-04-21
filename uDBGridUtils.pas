unit uDBGridUtils;

interface

uses
  System.SysUtils
  , System.Classes
  , Vcl.DBGrids
  , Data.DB
  , FireDAC.Comp.Client
  , Vcl.Dialogs;

type
  TDBGridUtils = class
  public
    class procedure OrdenarGrid(Column: TColumn);
    class procedure AjustarColunas(Grid: TDBGrid);
  end;

implementation

{ TDBGridUtils }

class procedure TDBGridUtils.OrdenarGrid(Column: TColumn);
var
  vDataset: TDataSet;
  vFieldName: string;
begin
  if not Assigned(Column.Field) then
    Exit;

  vDataset := Column.Field.DataSet;
  vFieldName := Column.FieldName;

  // Verificamos se o DataSet é um FDQuery ou FDMemTable (ambos usam IndexFieldNames)
  if (vDataset is TFDQuery) or (vDataset is TFDMemTable) then
  begin
    try
      vDataset.DisableControls;
      try
        // Lógica: Se já estiver ordenado ASC, muda para DESC (:D). Senão, ASC.
        if TFDQuery(vDataset).IndexFieldNames = vFieldName then
        begin
          TFDQuery(vDataset).IndexFieldNames := vFieldName + ':D';
        end
        else
        begin
          TFDQuery(vDataset).IndexFieldNames := vFieldName;
        end;

        vDataset.First;
      finally
        vDataset.EnableControls;
      end;
    except
      on E: Exception do
        ShowMessage('Erro ao ordenar: ' + E.Message);
    end;
  end;
end;

class procedure TDBGridUtils.AjustarColunas(Grid: TDBGrid);
var
  i: Integer;
  vDataSet: TDataSet;
begin
  vDataSet := Grid.DataSource.DataSet;
  if (not Assigned(vDataSet)) or (not vDataSet.Active) then
    Exit;

  // Ajusta a largura da coluna baseado no tamanho do DisplayLabel ou tamanho do campo
  for i := 0 to Grid.Columns.Count - 1 do
  begin
    if Assigned(Grid.Columns[i].Field) then
    begin
      // Define uma largura mínima baseada no título da coluna
      Grid.Columns[i].Width := (Length(Grid.Columns[i].Title.Caption) + 2) * 8;
    end;
  end;
end;

end.
