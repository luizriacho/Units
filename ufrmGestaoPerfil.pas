unit ufrmGestaoPerfil;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, 
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids, 
  Vcl.Buttons, FireDAC.Comp.Client;

type
  TfrmGestaoPerfil = class(TForm)
    lblPerfis: TLabel;
    cbPerfis: TComboBox;
    btnNovoPerfil: TBitBtn;
    StringGridPerm: TStringGrid;
    btnSalvar: TBitBtn;
    btnFechar: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure cbPerfisChange(Sender: TObject);
    procedure btnNovoPerfilClick(Sender: TObject);
    procedure btnSalvarClick(Sender: TObject);
    procedure btnFecharClick(Sender: TObject);
  private
    FConnection: TFDConnection;
    procedure CarregarPerfis;
    procedure CarregarModulosEPermissoes;
    function ObterIdPerfilSelecionado: Integer;
  public
    class procedure Exibir(AConnection: TFDConnection);
  end;

implementation

{$R *.dfm}

class procedure TfrmGestaoPerfil.Exibir(AConnection: TFDConnection);
var
  Frm: TfrmGestaoPerfil;
begin
  Frm := TfrmGestaoPerfil.Create(nil);
  try
    Frm.FConnection := AConnection;
    Frm.ShowModal;
  finally
    Frm.Free;
  end;
end;

procedure TfrmGestaoPerfil.FormShow(Sender: TObject);
begin
  StringGridPerm.Cells[0, 0] := 'Módulo';
  StringGridPerm.Cells[1, 0] := 'Acessar (S/N)';
  StringGridPerm.Cells[2, 0] := 'Inserir (S/N)';
  StringGridPerm.Cells[3, 0] := 'Editar (S/N)';
  StringGridPerm.Cells[4, 0] := 'Excluir (S/N)';

  StringGridPerm.ColWidths[0] := 200;
  StringGridPerm.ColWidths[1] := 90;
  StringGridPerm.ColWidths[2] := 90;
  StringGridPerm.ColWidths[3] := 90;
  StringGridPerm.ColWidths[4] := 90;

  CarregarPerfis;
end;

procedure TfrmGestaoPerfil.CarregarPerfis;
var
  Qry: TFDQuery;
begin
  cbPerfis.Items.Clear;
  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := FConnection;
    Qry.SQL.Text := 'SELECT ID_PERFIL, NOME_PERFIL FROM PERFIL ORDER BY NOME_PERFIL';
    Qry.Open;

    while not Qry.Eof do
    begin
      cbPerfis.Items.AddObject(Qry.FieldByName('NOME_PERFIL').AsString, TObject(Qry.FieldByName('ID_PERFIL').AsInteger));
      Qry.Next;
    end;

    if cbPerfis.Items.Count > 0 then
    begin
      cbPerfis.ItemIndex := 0;
      CarregarModulosEPermissoes;
    end;
  finally
    Qry.Free;
  end;
end;

function TfrmGestaoPerfil.ObterIdPerfilSelecionado: Integer;
begin
  Result := 0;
  if cbPerfis.ItemIndex >= 0 then
    Result := Integer(cbPerfis.Items.Objects[cbPerfis.ItemIndex]);
end;

procedure TfrmGestaoPerfil.CarregarModulosEPermissoes;
var
  Qry: TFDQuery;
  IdPerfil, Linha: Integer;
begin
  IdPerfil := ObterIdPerfilSelecionado;
  if IdPerfil = 0 then Exit;

  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := FConnection;
    Qry.SQL.Text := 'SELECT M.ID_MODULO, M.NOME_MODULO, ' +
                    'COALESCE(P.CAN_ACCESS, ''N'') AS CAN_ACCESS, ' +
                    'COALESCE(P.CAN_INSERT, ''N'') AS CAN_INSERT, ' +
                    'COALESCE(P.CAN_EDIT, ''N'') AS CAN_EDIT, ' +
                    'COALESCE(P.CAN_DELETE, ''N'') AS CAN_DELETE ' +
                    'FROM MODULO M ' +
                    'LEFT JOIN PERMISSAO_PERFIL P ON (P.ID_MODULO = M.ID_MODULO AND P.ID_PERFIL = :ID_PERFIL) ' +
                    'ORDER BY M.NOME_MODULO';
    Qry.ParamByName('ID_PERFIL').AsInteger := IdPerfil;
    Qry.Open;

    StringGridPerm.RowCount := Qry.RecordCount + 1;
    Linha := 1;

    while not Qry.Eof do
    begin
      StringGridPerm.Cells[0, Linha] := Qry.FieldByName('NOME_MODULO').AsString;
      StringGridPerm.Cells[1, Linha] := Qry.FieldByName('CAN_ACCESS').AsString;
      StringGridPerm.Cells[2, Linha] := Qry.FieldByName('CAN_INSERT').AsString;
      StringGridPerm.Cells[3, Linha] := Qry.FieldByName('CAN_EDIT').AsString;
      StringGridPerm.Cells[4, Linha] := Qry.FieldByName('CAN_DELETE').AsString;

      Inc(Linha);
      Qry.Next;
    end;
  finally
    Qry.Free;
  end;
end;

procedure TfrmGestaoPerfil.cbPerfisChange(Sender: TObject);
begin
  CarregarModulosEPermissoes;
end;

procedure TfrmGestaoPerfil.btnNovoPerfilClick(Sender: TObject);
var
  NomeNovoPerfil: string;
  Qry: TFDQuery;
begin
  NomeNovoPerfil := InputBox('Novo Perfil', 'Informe o nome do novo perfil:', '');
  if Trim(NomeNovoPerfil) = '' then Exit;

  Qry := TFDQuery.Create(nil);
  try
    Qry.Connection := FConnection;
    Qry.SQL.Text := 'INSERT INTO PERFIL (NOME_PERFIL) VALUES (:NOME)';
    Qry.ParamByName('NOME').AsString := UpperCase(Trim(NomeNovoPerfil));
    Qry.ExecSQL;

    CarregarPerfis;
    cbPerfis.ItemIndex := cbPerfis.Items.IndexOf(UpperCase(Trim(NomeNovoPerfil)));
    CarregarModulosEPermissoes;
  finally
    Qry.Free;
  end;
end;

procedure TfrmGestaoPerfil.btnFecharClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfrmGestaoPerfil.btnSalvarClick(Sender: TObject);
var
  Qry, QryId: TFDQuery;
  IdPerfil, IdModulo, I: Integer;
  NomeModulo: string;
begin
  IdPerfil := ObterIdPerfilSelecionado;
  if IdPerfil = 0 then Exit;

  Qry := TFDQuery.Create(nil);
  QryId := TFDQuery.Create(nil);
  try
    Qry.Connection := FConnection;
    QryId.Connection := FConnection;

    for I := 1 to StringGridPerm.RowCount - 1 do
    begin
      NomeModulo := StringGridPerm.Cells[0, I];
      if Trim(NomeModulo) = '' then Continue;

      // Busca o ID do módulo
      QryId.Close;
      QryId.SQL.Text := 'SELECT ID_MODULO FROM MODULO WHERE UPPER(NOME_MODULO) = UPPER(:NOME)';
      QryId.ParamByName('NOME').AsString := NomeModulo;
      QryId.Open;

      if not QryId.IsEmpty then
      begin
        IdModulo := QryId.FieldByName('ID_MODULO').AsInteger;

        Qry.Close;
        Qry.SQL.Text := 'UPDATE OR INSERT INTO PERMISSAO_PERFIL ' +
                        '(ID_PERFIL, ID_MODULO, CAN_ACCESS, CAN_INSERT, CAN_EDIT, CAN_DELETE) ' +
                        'VALUES (:ID_PERFIL, :ID_MODULO, :CAN_ACCESS, :CAN_INSERT, :CAN_EDIT, :CAN_DELETE) ' +
                        'MATCHING (ID_PERFIL, ID_MODULO)';
        Qry.ParamByName('ID_PERFIL').AsInteger := IdPerfil;
        Qry.ParamByName('ID_MODULO').AsInteger := IdModulo;
        Qry.ParamByName('CAN_ACCESS').AsString := UpperCase(StringGridPerm.Cells[1, I]);
        Qry.ParamByName('CAN_INSERT').AsString := UpperCase(StringGridPerm.Cells[2, I]);
        Qry.ParamByName('CAN_EDIT').AsString   := UpperCase(StringGridPerm.Cells[3, I]);
        Qry.ParamByName('CAN_DELETE').AsString := UpperCase(StringGridPerm.Cells[4, I]);
        Qry.ExecSQL;
      end;
    end;

    ShowMessage('Permissões do perfil salvas com sucesso!');
  finally
    Qry.Free;
    QryId.Free;
  end;
end;

end.