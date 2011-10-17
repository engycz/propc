unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo11 = class(TRttiItemServer)
  private
    FIntegerData: array[0..4] of Integer;
    FDoubleData: array[0..4] of Double;
    FStringData: array[0..4] of String;
    function GetIntegerArray(i: Integer): Integer;
    function GetRealArray(i: Integer): Double;
    function GetStringArray(i: Integer): String;
    procedure SetIntegerArray(i: Integer; Value: Integer);
    procedure SetRealArray(i: Integer; Value: Double);
    procedure SetStringArray(i: Integer; const Value: String);
  protected
    procedure LoadRttiItems(Proxy: TObject); override;
  public
  published
  end;

implementation
uses
  prOpcError;

{ TDemo11 }


function TDemo11.GetIntegerArray(i: Integer): Integer;
begin
  Result:= FIntegerData[i]
end;

function TDemo11.GetRealArray(i: Integer): Double;
begin
  Result:= FDoubleData[i]
end;

function TDemo11.GetStringArray(i: Integer): String;
begin
  Result:= FStringData[i]
end;

procedure TDemo11.SetIntegerArray(i, Value: Integer);
begin
  FIntegerData[i]:= Value
end;

procedure TDemo11.SetRealArray(i: Integer; Value: Double);
begin
  FDoubleData[i]:= Value
end;

procedure TDemo11.SetStringArray(i: Integer; const Value: String);
begin
  FStringData[i]:= Value
end;

procedure TDemo11.LoadRttiItems(Proxy: TObject);
begin
  {it is essential to call the default implementation first. This will load
  any 'standard' rtti items}
  inherited LoadRttiItems(Proxy);

  {this array is defined with a syntax of asNone, so it will generate a single
   server item, IntegerData. This will be an array of 5 integers}
  DefineIntegerArrayProperty('IntegerData', 5, asNone, GetIntegerArray, SetIntegerArray);

  {this array is defined with a syntax of asComma, and will generate 5 server
   items thus:
   'DoubleData,0', 'DoubleData,1' .. 'DoubleData,4'}
  DefineRealArrayProperty('DoubleData', 5, asComma, GetRealArray, SetRealArray);

  {this array is defined with a syntax of asBrackets, and will generate 5 server
   items thus:
   'StringData[0]', 'StringData[1]' .. 'StringData[4]'}
  DefineStringArrayProperty('StringData', 5, asBrackets, GetStringArray, SetStringArray)
end;


const
  ServerGuid: TGUID = '{C12606C1-AA2F-48A2-9659-4E46F9EAA6B2}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Rtti Array Demo';
  ServerVendor = 'Production Robots Eng. Ltd.';


initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo11.Create)
end.
 