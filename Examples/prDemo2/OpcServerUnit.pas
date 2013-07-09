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
  TDemo2 = class(TOpcItemServer)
  private
  protected
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

const
  TickCountHandle = 1;
  TimeOfDayHandle = 2;
  FormWidthHandle = 3;
  FormHeightHandle = 4;

{ TDemo2 }

function TDemo2.GetItemInfo(const ItemID: String; var AccessPath: string;
       var AccessRights: TAccessRights): Integer;
begin
  {Return a handle that will subsequently identify ItemID}
  {raise exception of type EOpcError if Item ID not recognised}
  if SameText(ItemId, 'TickCount') then
  begin
    Result:= TickCountHandle;
    AccessRights:= [iaRead]
  end else
  if SameText(ItemId, 'TimeOfDay') then
  begin
    Result:= TimeOfDayHandle;
    AccessRights:= [iaRead]
  end else
  if SameText(ItemId, 'FormWidth') then
  begin
    Result:= FormWidthHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  if SameText(ItemId, 'FormHeight') then
  begin
    Result:= FormHeightHandle;
    AccessRights:= [iaRead, iaWrite]
  end else
  begin
    raise EOpcError.Create(OPC_E_INVALIDITEMID)
  end
end;

procedure TDemo2.ReleaseHandle(ItemHandle: TItemHandle);
begin
  {Release the handle previously returned by GetItemInfo}
end;

procedure TDemo2.ListItemIds(List: TItemIDList);
begin
  {Call List.AddItemId(ItemId, AccessRights, VarType) for each ItemId}
  List.AddItemId('TickCount', [iaRead], varInteger);
  List.AddItemId('TimeOfDay', [iaRead], varOleStr);
  List.AddItemId('FormWidth', [iaRead, iaWrite], varInteger);
  List.AddItemId('FormHeight', [iaRead, iaWrite], varInteger);
end;

function TDemo2.GetItemValue(ItemHandle: TItemHandle;
                           var Quality: Word): OleVariant;
begin
  {return the value of the item identified by ItemHandle}
  case ItemHandle of
    TickCountHandle: Result:= Integer(Windows.GetTickCount);
    TimeOfDayHandle: Result:= TimeToStr(Now);
    FormWidthHandle: Result:= MainForm.Width;
    FormHeightHandle: Result:= MainForm.Height;
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

procedure TDemo2.SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant);
begin
  {set the value of the item identified by ItemHandle}
  case ItemHandle of
    FormWidthHandle: MainForm.Width:= Value;
    FormHeightHandle: MainForm.Height:= Value;
  else
    raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end
end;

const
  ServerGuid: TGUID = '{CAE8D0E1-117B-11D5-944B-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Simple demo';
  ServerVendor = 'Production Robots Eng. Ltd';


initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo2.Create)
end.

