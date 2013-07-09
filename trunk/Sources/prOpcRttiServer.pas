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
{History
v 1.15c 17/10/03
Do not return EUInfo for Boolean types cf 1.15.3.2

v 1.14 04/04/02 New Unit
}
unit prOpcRttiServer;
{$I prOpcCompilerDirectives.inc}
interface
uses
  Windows, Messages, SysUtils, Classes, TypInfo,
  prOpcComn, prOpcDa, prOpcError, prOpcTypes, prOpcServer;

type
  PRttiHandle = ^TRttiHandle;
  TRttiHandle = record
    case IsArray: Boolean of
      False: (Obj: TObject;
              PropInfo: PPropInfo);
      True: (ArrayProp: PArrayProperty);
  end;

  TBooleanArrayGetProc = function(i: Integer): Boolean of object;
  TIntegerArrayGetProc = function(i: Integer): Integer of object;
  TRealArrayGetProc = function(i: Integer): Double of object;
  TDateTimeArrayGetProc = function(i: Integer): TDateTime of object;
  TStringArrayGetProc = function(i: Integer): String of object;

  TBooleanArraySetProc = procedure(i: Integer; Value: Boolean) of object;
  TIntegerArraySetProc = procedure(i: Integer; Value: Integer) of object;
  TRealArraySetProc = procedure(i: Integer; Value: Double) of object;
  TDateTimeArraySetProc = procedure(i: Integer; Value: TDateTime) of object;
  TStringArraySetProc = procedure(i: Integer; const Value: String) of object;

  TRttiItemServer = class(TOpcItemServer)
  private
    FRttiItems: TStringList;
    FRttiProxy: TObject;
    procedure SetRttiProxy(Value: TObject);
    procedure DefineArrayProperty(ElementType: Integer; const Name: String;
      Size: Integer; Syntax: TArraySyntax;
      GetProc, SetProc: TMethod);
  protected
    function IsRecursive: Boolean; virtual;
    function GetExtendedItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights;
                        var EUInfo: IEUInfo;
                        var ItemProperties: IItemProperties): Integer; override;

    function GetItemValue(ItemHandle: TItemHandle;
                          var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
    procedure ListItemIDs(List: TItemIDList); override;

      {array properties}
    procedure DefineBooleanArrayProperty(const Name: String;
      Size: Integer; Syntax: TArraySyntax;
      GetProc: TBooleanArrayGetProc; SetProc: TBooleanArraySetProc);
    procedure DefineIntegerArrayProperty(const Name: String;
      Size: Integer; Syntax: TArraySyntax;
      GetProc: TIntegerArrayGetProc; SetProc: TIntegerArraySetProc);
    procedure DefineRealArrayProperty(const Name: String;
      Size: Integer; Syntax: TArraySyntax;
      GetProc: TRealArrayGetProc; SetProc: TRealArraySetProc);
    procedure DefineDateTimeArrayProperty(const Name: String;
      Size: Integer; Syntax: TArraySyntax;
      GetProc: TDateTimeArrayGetProc; SetProc: TDateTimeArraySetProc);
    procedure DefineStringArrayProperty(const Name: String;
      Size: Integer; Syntax: TArraySyntax;
      GetProc: TStringArrayGetProc; SetProc: TStringArraySetProc);

    procedure LoadRttiItems(Proxy: TObject); virtual;
  public
    constructor Create;
    destructor Destroy; override;
    property RttiProxy: TObject read FRttiProxy write SetRttiProxy;
  end;


implementation
uses
{$IFDEF D6UP}
  Variants,
{$ENDIF}
  ActiveX, prOpcVarUtils, prOpcClasses;

resourcestring
  SCannotChangeProxyWithItems = 'Cannot change rtti proxy while items exist';
  SNameserverDoesNotSupportRtti = 'Server does not support Rtti';
  SMustCallLoadRttiItems = 'LoadRttiItems must be called before DefineArrayProperty';
  SItemNameAlreadyDefined = 'Item name %s already defined';
  SCannotFindMethod = 'Cannot find method %s';
  SUnsupportedArrayType = 'Array type %d is not supported';
  SSyntaxConflict = 'Cannot use comma for array syntax and path separator'; {cf 1.14.18}

const
{$IFDEF F_UString}
  TypeStrings = [tkString, tkLString, tkWString, tkUString];
{$ELSE}
  TypeStrings = [tkString, tkLString, tkWString];
{$ENDIF}
  ValidTypeKinds = [tkInteger, tkEnumeration, tkFloat] + TypeStrings;

type
  TRttiItemList = class(TStringList)
    procedure FreeObjects;
  public
    procedure Clear; override;
    constructor Create;
    destructor Destroy; override;
  end;

{ Rtti utils }

function RttiHandleToVarType(ph: PRttiHandle): Integer;
begin
  if ph^.IsArray then
  begin
    Result:= ph^.ArrayProp^._Type
  end else
  with ph^.PropInfo^ do
  begin
    case PropType^^.Kind of
      tkInteger:
        Result:= VT_I4;
      tkEnumeration:
      begin
        if GetTypeData(PropType^)^.BaseType^ = TypeInfo(Boolean) then
          Result:= VT_BOOL
        else
          Result:= VT_I4
      end;
      tkFloat:
      begin
        if PropType^ = TypeInfo(TDateTime) then
          Result:= VT_DATE
        else
          Result:= VT_R8
      end;
{$IFDEF F_UString}
        tkString, tkLString, tkWString, tkUString:
{$ELSE}
        tkString, tkLString, tkWString:
{$ENDIF}
        Result:= VT_BSTR;
    else
      Result:= VT_EMPTY;
    end
  end
end;

function RttiHandleToAccessRights(ph: PRttiHandle): TAccessRights;
begin
  Result:= [];
  if ph^.IsArray then
  begin
    if Assigned(ph^.ArrayProp^.GetProc.Code) then
      Include(Result, iaRead);
    if Assigned(ph^.ArrayProp^.SetProc.Code) then
      Include(Result, iaWrite)
  end else
  begin
    if Assigned(ph^.PropInfo^.GetProc) then
      Include(Result, iaRead);
    if Assigned(ph^.PropInfo^.SetProc) then
      Include(Result, iaWrite)
  end
end;


type
  {this is VCL autostyle - ie everything by ref}
  TBoolSet = procedure(Self: TObject; i: Integer; const Val: Boolean);
  TIntSet = procedure(Self: TObject; i: Integer; const Val: Integer);
  TRealSet = procedure(Self: TObject; i: Integer; const Val: Double);
  TStringSet = procedure(Self: TObject; i: Integer; const Val: String);

  TBoolGet = function(Self: TObject; i: Integer): Boolean;
  TIntGet = function(Self: TObject; i: Integer): Integer;
  TRealGet = function(Self: TObject; i: Integer): Double;
  TStringGet = function(Self: TObject; i: Integer): String;

  {These types are also declared in prVarUtils, but I don't want to
  make them public}
  PWordBoolArray = ^TWordBoolArray;
  TWordBoolArray = array[Word] of WordBool;

  PIntegerArray = ^TIntegerArray;
  TIntegerArray = array[Word] of Integer;

  PRealArray = ^TRealArray;
  TRealArray = array[Word] of Double;

  PWideStringArray = ^TWideStringArray;
  TWideStringArray = array[Word] of WideString;

function TRttiItemServer.GetExtendedItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights;
                        var EUInfo: IEUInfo;
                        var ItemProperties: IItemProperties): Integer;
var
  ph: PRttiHandle absolute Result;
  i: Integer;
begin
  if Assigned(FRttiItems) and
     FRttiItems.Find(ItemID, i) then
  begin
    Result:= Integer(FRttiItems.Objects[i]);
    AccessRights:= RttiHandleToAccessRights(PRttiHandle(Result))
  end else
  begin
    raise EOpcError.Create(OPC_E_INVALIDITEMID)
  end;
  if not ph^.IsArray and
    (ph^.PropInfo^.PropType^ <> TypeInfo(Boolean)) and
    (ph^.PropInfo^.PropType^^.Kind = tkEnumeration) then   {cf 1.15.3.2}
    EUInfo:= TEnumeratedEUInfoFromRtti.Create(ph^.PropInfo^.PropType^)
end;

function TRttiItemServer.GetItemValue(ItemHandle: TItemHandle;
  var Quality: Word): OleVariant;
var
  ph: PRttiHandle absolute ItemHandle;
  Dest: Pointer;
  i: Integer;

begin
  with ph^ do
  if IsArray then
  begin
    with ArrayProp^ do
    if Index = -1 then
    begin
      Result:= VarArrayCreate([0, Size-1], _Type);
      Dest:= VarArrayLock(Result);
      try
        case _Type of
          VT_BOOL:
          for i:= 0 to Size - 1 do
            PWordBoolArray(Dest)^[i]:= TBooleanArrayGetProc(GetProc)(i);
          VT_I4:
          for i:= 0 to Size - 1 do
            PIntegerArray(Dest)^[i]:= TIntegerArrayGetProc(GetProc)(i);
          VT_R8, VT_DATE:
          for i:= 0 to Size - 1 do
            PRealArray(Dest)^[i]:= TRealArrayGetProc(GetProc)(i);
          VT_BSTR:
          for i:= 0 to Size - 1 do
            PWideStringArray(Dest)^[i]:= TStringArrayGetProc(GetProc)(i);
        end
      finally
        VarArrayUnlock(Result)
      end;
    end else
    begin
      case _Type of
        VT_BOOL: Result:= TBooleanArrayGetProc(GetProc)(Index);
        VT_I4: Result:= TIntegerArrayGetProc(GetProc)(Index);
        VT_R8: Result:= TRealArrayGetProc(GetProc)(Index);
        VT_DATE: VariantChangeType(Result,
          TRealArrayGetProc(GetProc)(Index), 0, VT_DATE);
        VT_BSTR: Result:= TStringArrayGetProc(GetProc)(Index);
      end
    end
  end else
  begin
    case RttiHandleToVarType(ph) of
      VT_BOOL: Result:= Boolean(GetOrdProp(Obj, PropInfo));
      VT_I4: Result:= GetOrdProp(Obj, PropInfo);
      VT_R8: Result:= GetFloatProp(Obj, PropInfo);
      VT_DATE: VariantChangeType(Result, GetFloatProp(Obj, PropInfo), 0, VT_DATE);
      VT_BSTR: Result:= GetStrProp(Obj, PropInfo);
    else
      raise EOpcError.Create(E_FAIL)
        {should never happen}
    end
  end
end;

procedure TRttiItemServer.ListItemIDs(List: TItemIDList);
var
  i: Integer;
  ph: PRttiHandle;
begin
  if Assigned(FRttiItems) then
  for i:= 0 to FRttiItems.Count - 1 do
  begin
    ph:= PRttiHandle(FRttiItems.Objects[i]);
    begin
      List.AddItemId(FRttiItems[i],
        RttiHandleToAccessRights(ph),
        RttiHandleToVarType(ph))
    end
  end
end;

procedure TRttiItemServer.SetItemValue(ItemHandle: TItemHandle;
  const Value: OleVariant);
var
  ph: PRttiHandle absolute ItemHandle;
  B: Boolean;
  Str: String;
  Src: Pointer;
  i: Integer;
begin
  with ph^ do
  if IsArray then
  begin
    with ArrayProp^ do
    if Index = -1 then
    begin
      Src:= VarArrayLock(Value);
      try
        case _Type of
          VT_BOOL:
          for i:= 0 to Size - 1 do
          begin
            B:= PWordBoolArray(Src)^[i];
            TBooleanArraySetProc(SetProc)(i, B)
          end;
          VT_I4:
          for i:= 0 to Size - 1 do
            TIntegerArraySetProc(SetProc)(i, PIntegerArray(Src)^[i]);
          VT_R8:
          for i:= 0 to Size - 1 do
            TRealArraySetProc(SetProc)(i, PRealArray(Src)^[i]);
          VT_DATE:
          for i:= 0 to Size - 1 do
            TDateTimeArraySetProc(SetProc)(i, PRealArray(Src)^[i]);
          VT_BSTR:
          for i:= 0 to Size - 1 do
          begin
            Str:= PWideStringArray(Src)^[i];
            TStringArraySetProc(SetProc)(i, Str)
          end;
        end
      finally
        VarArrayUnlock(Value)
      end;
    end else
    begin
      case _Type of
        VT_BOOL:
        begin
          B:= Value;
          TBooleanArraySetProc(SetProc)(Index, B);
        end;
        VT_I4: TIntegerArraySetProc(SetProc)(Index, Value);
        VT_R8: TRealArraySetProc(SetProc)(Index, Value);
        VT_Date: TDateTimeArraySetProc(SetProc)(Index, Value);
        VT_BSTR:
        begin
          Str:= Value;
          TStringArraySetProc(SetProc)(Index, Str);
        end;
      end
    end
  end else
  begin
    case RttiHandleToVarType(ph) of
      VT_BOOL:
      begin
        B:= Value;
        SetOrdProp(Obj, PropInfo, Integer(B))
      end;
      VT_I4:
        SetOrdProp(Obj, PropInfo, Value);
      VT_R8, VT_DATE:
        SetFloatProp(Obj, PropInfo, Value);
      VT_BSTR:
      begin
        Str:= Value;
        SetStrProp(Obj, PropInfo, Str)
      end
    end
  end
end;

constructor TRttiItemServer.Create;
begin
  inherited Create;
  LoadRttiItems(Self)
end;

destructor TRttiItemServer.Destroy;
begin
  FRttiItems.Free;
  inherited Destroy
end;


procedure TRttiItemServer.DefineArrayProperty(ElementType: Integer; const Name: String;
      Size: Integer; Syntax: TArraySyntax; GetProc, SetProc: TMethod);

procedure NewRttiEntry( const aName: String;
                        Index: Integer);
var
  ph: PRttiHandle;
  i: Integer;
begin
  try
    i:= FRttiItems.Add(AName);
    New(ph);
    with ph^ do
    begin
      IsArray:= true;
      New(ArrayProp);
      ArrayProp^.Size:= Size;
      ArrayProp^._Type:= ElementType;
      ArrayProp^.Index:= Index;
      ArrayProp^.GetProc:= GetProc;
      ArrayProp^.SetProc:= SetProc
    end;
    FRttiItems.Objects[i]:= TObject(ph)
  except
   on EStringListError do
     raise EOpcServer.CreateFmt(SItemNameAlreadyDefined, [Name])
  end
end;

var
  i: Integer;

const
  SyntaxFmt: array[asComma..asBrackets] of String =
   ('%s,%d', '%s[%d]');

begin
  if not Assigned(FRttiItems) then
    raise EOpcServer.Create(SMustCallLoadRttiItems);
  if not (ElementType in [VT_BOOL, VT_I4, VT_R8, VT_DATE, VT_BSTR]) then
    raise EOpcServer.CreateResFmt(@SUnsupportedArrayType, [ElementType]);
  if HierarchicalBrowsing and (PathDelimiter = ',') and (Syntax = asComma) then
    raise EOpcServer.CreateRes(@SSyntaxConflict);  {cf 1.14.18}
  if Syntax = asNone then
    NewRttiEntry(Name, -1)
  else
  for i:= 0 to Size - 1 do
    NewRttiEntry(Format(SyntaxFmt[Syntax], [Name, i]), i)
end;

procedure TRttiItemServer.LoadRttiItems(Proxy: TObject);
var
  Recursive: Boolean;

procedure AddObjProperties(const Name: String; Obj: TObject);
var
  ti: PTypeInfo;
  PropCount: Integer;
  PropList: PPropList;
  pi: PPropInfo;
  i: Integer;
  Kinds: TTypeKinds;

function NewName: String;
begin
  if Name = '' then
    Result:= pi^.Name
  else
    Result:= Name + PathDelimiter + pi^.Name
end;

var
  NewItem: PRttiHandle;

begin
  if Assigned(Obj) then
  begin
    ti:= Obj.ClassInfo;
    PropCount:= GetTypeData(ti)^.PropCount;
    if PropCount > 0 then
    begin
      GetMem(PropList, SizeOf(Pointer)*PropCount);
      try
        Kinds:= ValidTypeKinds;
        if Recursive then
          Include(Kinds, tkClass);
        PropCount:= GetPropList(ti, Kinds, PropList);
        for i:= 0 to PropCount - 1 do
        begin
          pi:= PropList^[i];
          if pi^.PropType^^.Kind = tkClass then
          begin
            AddObjProperties(NewName,
              GetObjectProp(Obj, pi))
          end else
          begin
            New(NewItem);
            NewItem^.Obj:= Obj;
            with NewItem^ do
            begin
              IsArray:= false;
              PropInfo:= pi;
            end;
            FRttiItems.AddObject(NewName, TObject(NewItem))
          end
        end
      finally
        FreeMem(PropList)
      end
    end
  end
end;

begin
  if not Assigned(FRttiItems) then
     FRttiItems:= TRttiItemList.Create;
  FRttiItems.Clear;
  Recursive:= IsRecursive;
  AddObjProperties('', Proxy)
end;

procedure TRttiItemServer.SetRttiProxy(Value: TObject);
var
  ItemCount: Integer;
begin
  with ItemList do
  begin
    ItemCount:= LockList.Count;
    UnlockList
  end;
  if ItemCount > 0 then
    raise EOpcServer.Create(SCannotChangeProxyWithItems);
  FRttiProxy:= Value;
  if Assigned(FRttiProxy) then
    LoadRttiItems(FRttiProxy)
  else
    LoadRttiItems(Self)
end;

procedure TRttiItemServer.DefineBooleanArrayProperty(const Name: String;
  Size: Integer; Syntax: TArraySyntax; GetProc: TBooleanArrayGetProc;
  SetProc: TBooleanArraySetProc);
begin
  DefineArrayProperty(VT_BOOL, Name, Size, Syntax,
    TMethod(GetProc), TMethod(SetProc))
end;

procedure TRttiItemServer.DefineIntegerArrayProperty(const Name: String;
  Size: Integer; Syntax: TArraySyntax; GetProc: TIntegerArrayGetProc;
  SetProc: TIntegerArraySetProc);
begin
  DefineArrayProperty(VT_I4, Name, Size, Syntax,
    TMethod(GetProc), TMethod(SetProc))
end;

procedure TRttiItemServer.DefineRealArrayProperty(const Name: String;
  Size: Integer; Syntax: TArraySyntax; GetProc: TRealArrayGetProc;
  SetProc: TRealArraySetProc);
begin
  DefineArrayProperty(VT_R8, Name, Size, Syntax,
    TMethod(GetProc), TMethod(SetProc))
end;

procedure TRttiItemServer.DefineStringArrayProperty(const Name: String;
  Size: Integer; Syntax: TArraySyntax; GetProc: TStringArrayGetProc;
  SetProc: TStringArraySetProc);
begin
  DefineArrayProperty(VT_BSTR, Name, Size, Syntax,
    TMethod(GetProc), TMethod(SetProc))
end;

procedure TRttiItemServer.DefineDateTimeArrayProperty(const Name: String;
  Size: Integer; Syntax: TArraySyntax; GetProc: TDateTimeArrayGetProc;
  SetProc: TDateTimeArraySetProc);
begin
  DefineArrayProperty(VT_DATE, Name, Size, Syntax,
    TMethod(GetProc), TMethod(SetProc))
end;

{ TRttiItemList }

procedure TRttiItemList.Clear;
begin
  FreeObjects;
  inherited Clear;
end;

constructor TRttiItemList.Create;
begin
  inherited Create;
  Sorted:= true;
  Duplicates:= dupError
end;

destructor TRttiItemList.Destroy;
begin
  FreeObjects;
  inherited Destroy
end;

procedure TRttiItemList.FreeObjects;
var
  i: Integer;
  ph: PRttiHandle;
begin
  for i:= 0 to Count -1 do
  begin
    ph:= PRttiHandle(Objects[i]);
    if ph^.IsArray then
      Dispose(ph^.ArrayProp);
    Dispose(ph)
  end;
end;

function TRttiItemServer.IsRecursive: Boolean;
begin
  Result:= true
end;


end.
