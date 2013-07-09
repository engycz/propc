{------------------------------------------------------------}
{The MIT License (MIT)

 prOpc Toolkit
 Copyright (c) 2000, 2001 Production Robots Engineering Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.}
{------------------------------------------------------------}
unit prOpcClasses;
{$I prOpcCompilerDirectives.Inc}
interface
uses
  Classes, TypInfo, ActiveX, prOpcTypes, prOpcServer;

type
  {abstract base class do not instantiate}
  TAnalogEU = class(TInterfacedObject, IEUInfo)
  private
    function EUType: TEuType;
    function EUInfo: OleVariant;
  protected
    procedure GetLimits(var Low, High: Double); virtual; abstract;
  end;

  {abstract base class do not instantiate}
  TEnumeratedEU = class(TInterfacedObject, IEUInfo)
  private
    function EUType: TEuType;
    function EUInfo: OleVariant;
  protected
    procedure GetEnumeratedNames(Names: TStrings); virtual; abstract;
  end;

  { Property Support }
  {abstract base class do not instantiate}
  TItemProperty = class(TInterfacedObject, IItemProperty)
  private
    FPid: LongWord;
  public
    constructor Create(aPid: Integer);
    function Description: string; virtual; abstract;
    function DataType: Integer; virtual; abstract;
    function GetPropertyValue: OleVariant; virtual; abstract;
    function Pid: LongWord;
  end;

  TItemProperties = class(TInterfacedObject, IItemProperties)
  private
    FList: TInterfaceList;
  protected
  public
    function GetPropertyItem(Index: Integer): IItemProperty;
    function GetProperty(Pid: LongWord): IItemProperty;
    procedure Add(const ItemProperty: IItemProperty);
    function Count: Integer;
    constructor Create;
    destructor Destroy; override;
  end;

  {basic fixed analog limits. Much more sophistication than this is possible
  by deriving your own class}
  TAnalogLimits = class(TAnalogEU)
  private
    FLow, FHigh: Double;
  protected
    procedure GetLimits(var Low, High: Double); override;
  public
    constructor Create(const aLow, aHigh: Double);
  end;

  {for use with enumerated rtti types}
  TEnumeratedEUInfoFromRtti = class(TEnumeratedEU)
  private
    FTypeInfo: PTypeInfo;
  protected
    procedure GetEnumeratedNames(Names: TStrings); override;
  public
    constructor Create(aTypeInfo: PTypeInfo);
  end;

  TEnumeratedEUInfoFromStrings = class(TEnumeratedEU, IStringsAdapter)
  private
    FStrings: TStrings;
    procedure ReferenceStrings(S: TStrings);
    procedure ReleaseStrings;
  protected
    procedure GetEnumeratedNames(Names: TStrings); override;
  public
    constructor Create(aStrings: TStrings);
  end;

implementation
uses
  {$IFDEF D6UP}
  Variants,
  {$ENDIF}
  prOpcDa;


resourcestring
  SPropertyAlreadyExists = 'Property %s with Pid %d already exists';

{ TItemProperties }

{ this class is inefficiently implemented - should be a sorted list
  keyed on PID &&& For small property lists this is probably not
  very important }

constructor TItemProperty.Create(aPid: Integer);
begin
  inherited Create;
  FPid:= aPid
end;

procedure TItemProperties.Add(const ItemProperty: IItemProperty);
var
  Prop: IItemProperty;
begin
  Prop:= GetProperty(ItemProperty.Pid);
  if Prop <> nil then
    raise EOpcError.CreateResFmt(@SPropertyAlreadyExists,
      [Prop.Description, Prop.Pid]);
  FList.Add(ItemProperty)
end;

function TItemProperties.Count: Integer;
begin
  Result:= FList.Count
end;

constructor TItemProperties.Create;
begin
  inherited Create;
  FList:= TInterfaceList.Create
end;

destructor TItemProperties.Destroy;
begin
  FList.Free;
  inherited Destroy
end;

function TItemProperties.GetPropertyItem(Index: Integer): IItemProperty;
begin
  Result := IItemProperty(FList[Index]);
end;

function TItemProperties.GetProperty(Pid: LongWord): IItemProperty;
var
  i: Integer;
  Prop: IItemProperty;
begin
  Result:= nil;
  for i:= 0 to FList.Count - 1 do
  begin
    Prop:= FList[i] as IItemProperty;
    if Prop.Pid = Pid then
    begin
      Result:= Prop;
      break
    end
  end
end;


{ TEUInfo }

function EUType: TOleEnum;
begin
  Result:= OPC_NOENUM
end;

{ TAnalogEU }

function TAnalogEU.EUInfo: OleVariant;
var
  h, l: Double;
begin
  GetLimits(l, h);
  Result:= VarArrayCreate([0,1], VT_R8);
  Result[0]:= l;
  Result[1]:= h
end;

function TAnalogEU.EUType: TEuType;
begin
  Result:= euAnalog
end;

{ TEnumeratedEU }

function TEnumeratedEU.EUInfo: OleVariant;
var
  s: TStringList;
  i: Integer;
begin
  s:= TStringList.Create;
  try
    GetEnumeratedNames(s);
    if s.Count = 0 then
    begin
      Result:= Unassigned
    end else
    begin
      Result:= VarArrayCreate([0, s.Count-1], VT_BSTR);
      for i:= 0 to s.Count - 1 do
        Result[i]:= s[i]
    end
  finally
    s.Free
  end
end;

function TEnumeratedEU.EUType: TEuType;
begin
  Result:= euEnumerated
end;



{ TEnumeratedTypeEUInfo }

constructor TEnumeratedEUInfoFromRtti.Create(aTypeInfo: PTypeInfo);
begin
  inherited Create;
  FTypeInfo:= aTypeInfo  {it is probably OK to keep a pointer to this
                          because Rtti is static global (it appears)} 
end;

procedure TEnumeratedEUInfoFromRtti.GetEnumeratedNames(Names: TStrings);
var
  i: Integer;
begin
  with GetTypeData(FTypeInfo)^ do
  for i:= MinValue to MaxValue do
    Names.Add(GetEnumName(FTypeInfo, i))
end;


{ TAnalogLimits }

constructor TAnalogLimits.Create(const aLow, aHigh: Double);
begin
  inherited Create;
  FLow:= aLow;
  FHigh:= aHigh
end;

procedure TAnalogLimits.GetLimits(var Low, High: Double);
begin
  Low:= FLow;
  High:= FHigh
end;



function TItemProperty.Pid: LongWord;
begin
  Result:= FPid
end;

{ TEnumeratedEUInfoFromStrings }

constructor TEnumeratedEUInfoFromStrings.Create(aStrings: TStrings);
begin
  inherited Create;
  if Assigned(aStrings) then
    aStrings.StringsAdapter:= Self
end;

procedure TEnumeratedEUInfoFromStrings.GetEnumeratedNames(Names: TStrings);
begin
  if Assigned(FStrings) then
    Names.AddStrings(FStrings)
end;

procedure TEnumeratedEUInfoFromStrings.ReferenceStrings(S: TStrings);
begin
  FStrings:= S
end;

procedure TEnumeratedEUInfoFromStrings.ReleaseStrings;
begin
  FStrings:= nil
end;

end.
