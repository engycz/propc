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
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes, Forms;

type
  TDemo13 = class(TRttiItemServer)
  private
    procedure LogMsg(const Msg: string);
    function GetFormBorderStyle: TFormBorderStyle;
    procedure SetFormBorderStyle(const Value: TFormBorderStyle);
  protected
    procedure OnClientConnect(Client: TClientInfo); override;
    procedure OnClientDisconnect(Client: TClientInfo); override;
    procedure OnClientSetName(Client: TClientInfo); override;
    procedure OnAddGroup(Group: TGroupInfo); override;
    procedure OnRemoveGroup(Group: TGroupInfo); override;
    procedure OnAddItem(Item: TGroupItemInfo); override;
    procedure OnRemoveItem(Item: TGroupItemInfo); override;
  public
  published
    property BorderStyle: TFormBorderStyle read GetFormBorderStyle
       write SetFormBorderStyle;
  end;

implementation
uses
  prOpcError, Windows, MainUnit;

{ TDemo13 }

procedure TDemo13.OnClientConnect(Client: TClientInfo);
begin
  MainForm.ExitServerButton.Enabled:= false;
  LogMsg('Client connected Ok')
end;

procedure TDemo13.OnClientDisconnect(Client: TClientInfo);
begin
  if ClientCount = 0 then
    MainForm.ExitServerButton.Enabled:= true;
  LogMsg('Client disconnected Ok');
end;

procedure TDemo13.OnClientSetName(Client: TClientInfo);
begin
  LogMsg('Client SetName Ok')
end;

procedure TDemo13.OnAddGroup(Group: TGroupInfo);
begin
  LogMsg(Format('Add Group %s ok', [Group.Name]))
end;

procedure TDemo13.OnRemoveGroup(Group: TGroupInfo);
begin
  LogMsg(Format('Remove Group %s ok', [Group.Name]))
end;

procedure TDemo13.OnAddItem(Item: TGroupItemInfo);
begin
  LogMsg(Format('Add Item %s ok', [Item.ItemID]))
end;

procedure TDemo13.OnRemoveItem(Item: TGroupItemInfo);
begin
  LogMsg(Format('Remove Item %s ok', [Item.ItemID]))
end;

procedure TDemo13.LogMsg(const Msg: string);
begin
  if Assigned(MainForm) then
    MainForm.DebugLog.Items.Add(Msg)
end;

const
  ServerGuid: TGUID = '{DCE243DF-2003-4D5C-B78E-CDECA980EA03}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Enumerated Type demo';
  ServerVendor = 'Production Robots Eng. Ltd.';

function TDemo13.GetFormBorderStyle: TFormBorderStyle;
begin
  Result:= MainForm.BorderStyle
end;

procedure TDemo13.SetFormBorderStyle(const Value: TFormBorderStyle);
begin
  MainForm.BorderStyle:= Value
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo13.Create)
end.

