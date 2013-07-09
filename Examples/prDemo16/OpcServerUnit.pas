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
unit OpcServerUnit;

interface

uses
  SysUtils, Classes, prOpcServer, prOpcTypes;

type
  TDemo16 = class(TOpcItemServer)
  private
  protected
    function Options: TServerOptions; override;
  public
    function GetItemInfo(const ItemID: String; var AccessPath: string;
      var AccessRights: TAccessRights): Integer; override;
    procedure ReleaseHandle(ItemHandle: TItemHandle); override;
    procedure ListItemIds(List: TItemIDList); override;
    function GetItemValue(ItemHandle: TItemHandle;
                            var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
  end;

implementation
uses
  prOpcError, Windows, MainUnit;

{Demonstrate hierarchical browsing. This example is designed to show the
principles of defining a hierarchical namespace as clearly as possible. It is
not meant to be an example of how you should implement your server - I hope you
will be able to come up with something a lot more elegant than this}

{Test Namespace looks like this:

LINE1
	PLC
		INT_001: Integer
		INT_002: Integer
		STR_001: Integer
		STR_002: Integer
	WEIGHER
		WEIGHT: Double
		ERROR_STATUS: Integer
	WRAPPER
		STATUS: String
		LOW_FILM: Boolean


There is no formal relationship between the structure of the data and the
structure of the namespace. We will demonstrate this by defining the data
as a list of globals
}

var
  PlcInt1: Integer = 5;                 {LINE1.PLC.INT_001}
  PlcInt2: Integer = 6;                 {LINE1.PLC.INT_002}
  PlcStr1: WideString = 'Plc String 1'; {LINE1.PLC.STR_001}
  PlcStr2: WideString = 'Plc String 2'; {LINE1.PLC.STR_002}
  Weight: Double = 1.134;               {LINE1.WEIGHER.WEIGHT}
  WeigherStatus: Integer = 2;           {LINE1.WEIGHER.ERROR_STATUS}
  WrapperStatus:WideString = 'OK';      {LINE1.WRAPPER.STATUS}
  WrapperLow: Boolean = false;          {LINE1.WRAPPER.LOW_FILM}

{ TDemo16 }

{$IFDEF NewBranch}

procedure TDemo16.ListItemIDs(List: TItemIDList);
begin
  with List.AddBranch('LINE1') do
  begin
    with AddBranch('PLC') do
    begin
      AddItemId('INT_001', AllAccess, varInteger);
      AddItemId('INT_002', AllAccess, varInteger);
      AddItemId('STR_001', AllAccess, varInteger);
      AddItemId('STR_002', AllAccess, varInteger);
    end;
    with AddBranch('WEIGHER') do
    begin
      AddItemId('WEIGHT', AllAccess, varInteger);
      AddItemId('ERROR_STATUS', AllAccess, varInteger)
    end;
    with AddBranch('WRAPPER') do
    begin
      AddItemId('STATUS', AllAccess, varInteger);
      AddItemId('LOW_FILM', AllAccess, varInteger)
    end
  end
end;

{$ELSE}

procedure TDemo16.ListItemIDs(List: TItemIDList);
begin  {use qualified Ids}
  List.AddItemId('LINE1.PLC.INT_001', AllAccess, varInteger);
  List.AddItemId('LINE1.PLC.INT_002', AllAccess, varInteger);
  List.AddItemId('LINE1.PLC.STR_001', AllAccess, varOleStr);
  List.AddItemId('LINE1.PLC.STR_002', AllAccess, varOleStr);
  List.AddItemId('LINE1.WEIGHER.WEIGHT', AllAccess, varDouble);
  List.AddItemId('LINE1.WEIGHER.ERROR_STATUS', AllAccess, varInteger);
  List.AddItemId('LINE1.WRAPPER.STATUS', AllAccess, varOleStr);
  List.AddItemId('LINE1.WRAPPER.LOW_FILM', AllAccess, varBoolean)
end;

{$ENDIF}

function TDemo16.GetItemInfo(const ItemID: String; var AccessPath: string;
       var AccessRights: TAccessRights): Integer;
begin
  {Return a handle that will subsequently identify ItemID}
  {raise exception of type EOpcError if Item ID not recognised}
  if SameText(ItemID, 'LINE1.PLC.INT_001') then
    Result:= 0
  else
  if SameText(ItemID, 'LINE1.PLC.INT_002') then
    Result:= 1
  else
  if SameText(ItemID, 'LINE1.PLC.STR_001') then
    Result:= 2
  else
  if SameText(ItemID, 'LINE1.PLC.STR_002') then
    Result:= 3
  else
  if SameText(ItemID, 'LINE1.WEIGHER.WEIGHT') then
    Result:= 4
  else
  if SameText(ItemID, 'LINE1.WEIGHER.ERROR_STATUS') then
    Result:= 5
  else
  if SameText(ItemID, 'LINE1.WRAPPER.STATUS') then
    Result:= 6
  else
  if SameText(ItemID, 'LINE1.WRAPPER.LOW_FILM') then
    Result:= 7
  else
    raise EOpcError.Create(OPC_E_INVALIDITEMID)
end;

procedure TDemo16.ReleaseHandle(ItemHandle: TItemHandle);
begin
  {Release the handle previously returned by GetItemInfo}
end;

function TDemo16.GetItemValue(ItemHandle: TItemHandle;
                           var Quality: Word): OleVariant;
begin
  {return the value of the item identified by ItemHandle}
  case ItemHandle of
    0: Result:= PlcInt1;
    1: Result:= PlcInt2;
    2: Result:= PlcStr1;
    3: Result:= PlcStr2;
    4: Result:= Weight;
    5: Result:= WeigherStatus;
    6: Result:= WrapperStatus;
    7: Result:= WrapperLow;
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

procedure TDemo16.SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant);
begin
  {set the value of the item identified by ItemHandle}
  case ItemHandle of
    0: PlcInt1:= Value;
    1: PlcInt2:= Value;
    2: PlcStr1:= Value;
    3: PlcStr2:= Value;
    4: Weight:= Value;
    5: WeigherStatus:= Value;
    6: WrapperStatus:= Value;
    7: WrapperLow:= Value;
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

const
  ServerGuid: TGUID = '{8D0B5528-F4B5-4FEE-84F2-4DA6735678ED}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Hierarchical Demo';
  ServerVendor = 'Production Robots Eng. Ltd';


function TDemo16.Options: TServerOptions;
begin
  Result:= [soHierarchicalBrowsing, soAlwaysAllocateErrorArrays]
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo16.Create)
end.

