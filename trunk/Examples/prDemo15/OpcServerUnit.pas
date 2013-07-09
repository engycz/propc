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
{this example implements a standard server without rtti. Instead of
overriding GetItemInfo, the nameserver overrides GetExtendedItemInfo
in order to provide a EUInfo object. There are a number of standard
Objects available from the unit prOpcClasses which implement IEUInfo,
or you can derive your own, although it is not anticipated that this
will often be necessary}

interface

uses
  SysUtils, Classes, prOpcServer, prOpcTypes;

type
  TDemo15 = class(TOpcItemServer)
  private
  protected
    procedure OnItemValueChange(Item: TGroupItemInfo); override;
    function GetExtendedItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights;
                        var EUInfo: IEUInfo;
                        var ItemProperties: IItemProperties): Integer; override;
    procedure ListItemIDs(List: TItemIDList); override;
    function GetItemValue(ItemHandle: TItemHandle;
                          var Quality: Word): OleVariant; override;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;
  public
  published
  end;

implementation
uses
  prOpcError, Windows, MainUnit, StdCtrls, ComCtrls, TypInfo,
  prOpcClasses;  {for TAnalogEU}

{ TDemo15 }

function TDemo15.GetExtendedItemInfo(const ItemID: String;
  var AccessPath: String; var AccessRights: TAccessRights;
  var EUInfo: IEUInfo; var ItemProperties: IItemProperties): Integer;
begin
  if SameText('Trackbar', ItemId) then
  begin
    Result:= Integer(MainForm.TrackBar);
    EUInfo:= TAnalogLimits.Create(
      MainForm.TrackBar.Min, MainForm.TrackBar.Max)
  end else
  if SameText('Listbox', ItemId) then
  begin
    Result:= Integer(MainForm.ListBox);
    EUInfo:= TEnumeratedEUInfoFromStrings.Create(MainForm.ListBox.Items)
  end else
  begin
    raise EOpcError.Create(OPC_E_UNKNOWNITEMID);
  end
end;

function TDemo15.GetItemValue(ItemHandle: TItemHandle;
  var Quality: Word): OleVariant;
var
  Component: TComponent;
begin
  Component:= TComponent(ItemHandle);
  if Component is TTrackbar then
    Result:= TTrackbar(Component).Position
  else
  if Component is TComboBox then
    Result:= TComboBox(Component).ItemIndex
end;

procedure TDemo15.ListItemIDs(List: TItemIDList);
begin
  List.AddItemID('TrackBar', AllAccess, varInteger);
  List.AddItemID('ListBox', AllAccess, varInteger)
end;

procedure TDemo15.SetItemValue(ItemHandle: TItemHandle;
  const Value: OleVariant);

  procedure SetTrackBar(tb: TTrackBar; Val: Integer);
  begin
    if Val < tb.Min then
      tb.Position:= tb.Min
    else
    if Val > tb.Max then
      tb.Position:= tb.Max
    else
      tb.Position:= Val
  end;

  procedure SetCombo(cb: TComboBox; Val: Integer);
  begin
    if (Val < 0) or (Val >= cb.Items.Count) then
      cb.ItemIndex:= -1
    else
      cb.ItemIndex:= Val
  end;

var
  Component: TComponent;
begin
  Component:= TComponent(ItemHandle);
  if Component is TTrackbar then
    SetTrackBar(TTrackbar(Component), Value)
  else
  if Component is TComboBox then
    SetCombo(TComboBox(Component), Value)
end;

procedure TDemo15.OnItemValueChange(Item: TGroupItemInfo);
begin
  if TObject(Item.ItemHandle) = MainForm.TrackBar then
    MainForm.ShowDeadband(Item.LastUpdateValue,
      Item.Group.PercentDeadband)
end;

const
  ServerGuid: TGUID = '{455F0BC1-E2A1-4170-81AB-F8B910732AE9}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Deadband and Enumerates';
  ServerVendor = 'Production Robots Eng. Ltd.';

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo15.Create)
end.

