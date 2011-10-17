{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit prOpcVarUtils;
{$I prOpcCompilerDirectives.inc}

{History
1.10    New Unit


1.13.1  Bug in CompareVariant. Some Variants were not unlocked after compare.
        left locked. This bug became manifest under D6 but is present regardless
        of compiler. Only release 1.12 is affected.

1.16c Fixed bug in ReadVarArrayBoolean
        }
interface
uses
  SysUtils, Windows;

function CreateVarArray( const Data: array of Smallint): OleVariant; overload;
function CreateVarArray( const Data: array of Longint): OleVariant; overload;
function CreateVarArray( const Data: array of Single): OleVariant; overload;
function CreateVarArray( const Data: array of Double): OleVariant; overload;
function CreateVarArray( const Data: array of TDateTime): OleVariant; overload;
function CreateVarArray( const Data: array of WideString): OleVariant; overload;
function CreateVarArray( const Data: array of Boolean): OleVariant; overload;
function CreateVarArray( const Data: array of Shortint): OleVariant; overload;
function CreateVarArray( const Data: array of Byte): OleVariant; overload;
function CreateVarArray( const Data: array of Word): OleVariant; overload;
function CreateVarArray( const Data: array of Longword): OleVariant; overload;

procedure ReadVarArray( const Value: OleVariant; var Data);

function CompareVariant( const A, B: OleVariant): Boolean;
{the = operator does not support arrays. returns false if any are different}

procedure CheckArray(const Value: OleVariant);
{check that Value is valid array i.e
* IsArray
* ElementType in [VT_I2, VT_I4, VT_R4, VT_R8, VT_DATE, VT_BSTR, VT_BOOL,
                  VT_I1, VT_UI1, VT_UI2, VT_UI4, VT_INT, VT_UINT]

* Dimension = 1
* LowerBound = 0

raises EVariantError if not correct
}

function ArrayDataSize(const Value: OleVariant): Cardinal;

implementation
uses
{$IFDEF D6UP}
  Variants,
{$ENDIF}
  prOpcTypes, ActiveX;

resourcestring
  SNotArray = 'Variant is not an array';
  SUnsupportedType = 'Variant array has unsupported element type';
  SUnsupportedDim = 'Multi-dimensional variant arrays not supported';
  SUnsupportedBounds = 'Variant arrays must be zero-based';
  SIncorrectSize = 'Variant array is incorrect size';

type
  {types requiring element by element typechanges}
  PWordBoolArray = ^TWordBoolArray;
  TWordBoolArray = array[Word] of WordBool;

  PBooleanArray = ^TBooleanArray;
  TBooleanArray = array[Word] of Boolean;

{  PStringArray = ^TStringArray;
  TStringArray = array[Word] of String; }

  PWideStringArray = ^TWideStringArray;
  TWideStringArray = array[Word] of WideString;

function CreateVarArray( const Data: array of Smallint): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_I2);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Smallint));
  VarArrayUnlock(Result);
end;

function CreateVarArray( const Data: array of Longint): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_I4);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Longint));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Single): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_R4);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Single));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Double): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_R8);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Double));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of TDateTime): OleVariant; overload;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_DATE);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(TDateTime));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of WideString): OleVariant;
{need to do this to ensure that reference counts are correct on string}
var
  VData: PWideStringArray;
  i: Integer;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_BSTR);
  VData:= VarArrayLock(Result);
  for i:= 0 to High(Data) do
    VData^[i]:= Data[i];
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Boolean): OleVariant;
{Variants hold word bools so cannot just push data}
var
  VData: PWordBoolArray;
  i: Integer;
  B: Boolean;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_BOOL);
  VData:= VarArrayLock(Result);
  for i:= 0 to High(Data) do
  begin
    B:= Data[i]; {intermediate var stops internal compiler error}
    VData^[i]:= B;
  end;
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Shortint): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_I1);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Shortint));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Byte): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_UI1);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Byte));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Word): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_UI2);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Word));
  VarArrayUnlock(Result)
end;

function CreateVarArray( const Data: array of Longword): OleVariant;
begin
  Result:= VarArrayCreate([0, High(Data)], VT_UI4);
  Move(Data, VarArrayLock(Result)^, Length(Data)*SizeOf(Longword));
  VarArrayUnlock(Result)
end;

procedure ReadVarArrayWideString( const Value: OleVariant; var Data);
var
  VD: TVarData absolute Value;
  Dest: TWideStringArray absolute Data;
  Source: PWideStringArray;
  i: Integer;
begin
  Source:= VarArrayLock(Value);
  for i:= 0 to VD.VArray^.Bounds[0].ElementCount - 1 do
    Dest[i]:= Source^[i];
  VarArrayUnlock(Value)
end;

procedure ReadVarArrayBoolean( const Value: OleVariant; var Data);
var
  VD: TVarData absolute Value;
  Dest: TBooleanArray absolute Data;
  {Dest: TWordBoolArray absolute Data;  1.16c}
  Source: PWordBoolArray;
  i: Integer;
begin
  Source:= VarArrayLock(Value);
  for i:= 0 to VD.VArray^.Bounds[0].ElementCount - 1 do
    Dest[i]:= Source^[i];
  VarArrayUnlock(Value)
end;

procedure ReadVarArray( const Value: OleVariant; var Data);
begin
  CheckArray(Value);
  case VarType(Value) and not VT_ARRAY of
    VT_BSTR: ReadVarArrayWideString(Value, Data);
    VT_BOOL: ReadVarArrayBoolean(Value, Data);
  else
    begin
      Move(VarArrayLock(Value)^, Data, ArrayDataSize(Value));
      VarArrayUnlock(Value)
    end
  end
end;

procedure CheckArray(const Value: OleVariant);
var
  VD: TVarData absolute Value;
begin
  if (VD.VType and VT_ARRAY) = 0 then
    raise EVariantError.Create(SNotArray);
  if not IsSupportedVarType(VD.VType) then
    raise EVariantError.Create(SUnsupportedType);
  with VD.VArray^ do
  begin
    if DimCount <> 1 then
      raise EVariantError.Create(SUnsupportedDim);
    if Bounds[0].LowBound <> 0 then
      raise EVariantError.Create(SUnsupportedBounds)
  end
end;

function ArrayDataSize(const Value: OleVariant): Cardinal;
var
  VD: TVarData absolute Value;
begin
  with VD.VArray^ do
    Result:= ElementSize*Bounds[0].ElementCount
end;

function CompareVariant( const A, B: OleVariant): Boolean;
var
  vda: TVarData absolute A;
  vdb: TVarData absolute B;
  i: Integer;
  sta, stb: PWideStringArray;
  pa, pb: Pointer;
begin
  if vda.VType <> vdb.VType then
  begin
    Result:= false
  end else
  if (vda.VType and VT_ARRAY) = 0 then
  begin
    Result:= A = B
  end else
  if (vda.VArray^.DimCount = 1) and
     (vdb.VArray^.DimCount = 1) and
     (vda.VArray^.Bounds[0].ElementCount = vdb.VArray^.Bounds[0].ElementCount) and
     (vda.VArray^.ElementSize = vdb.VArray^.ElementSize) then
  begin
    if vda.VType = (VT_BSTR or VT_ARRAY) then
    begin
      sta:= VarArrayLock(A);
      stb:= VarArrayLock(B);
      try {cf 1.13.1}
        Result:= true;
        for i:= 0 to vda.VArray^.Bounds[0].ElementCount - 1 do
        if sta^[i] <> stb^[i] then
        begin
          Result:= false;
          break
        end
      finally
        VarArrayUnlock(A);
        VarArrayUnlock(B)
      end
    end else
    begin
      pa:= VarArrayLock(A);
      pb:= VarArrayLock(B);
      try
        Result:= CompareMem(pa, pb,
           vda.VArray^.Bounds[0].ElementCount*vda.VArray^.ElementSize)
      finally
        VarArrayUnlock(A);
        VarArrayUnlock(B)
      end
    end
  end else
  begin
    Result:= false
  end
end;

end.
