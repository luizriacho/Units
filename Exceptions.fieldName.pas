unit Exceptions.FieldName;

interface

uses
  system.SysUtils;

type
  ExceptionsFieldName = class(Exception)
  private
    FFieldName : string;
  public
    constructor Create(const AMessage, AFieldName: string); reintroduce;
    property FieldName: string read FFieldName write FFieldName;

  end;

implementation

constructor ExceptionsFieldName.Create(const AMessage, AFieldName: string);
begin
  self.Message := AMessage;
  FFieldname := AFieldName;
end;

end.
