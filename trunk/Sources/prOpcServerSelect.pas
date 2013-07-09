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
unit prOpcServerSelect;
{$I prOpcCompilerDirectives.inc}
{History

Release 1.15c Oct 2003
1.15.20 added search options to TServerSelectDlg  cf 1.15.3.1

Release 1.14 April 2002
1.14.1 Changed some string literals to resourcestring.
}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  prOpcClient, ComCtrls, StdCtrls, prOpcTypes, prOpcEnum,
  Buttons, prOpcBrowser;

type
  TServerSelectDlg = class(TForm)
    List: TListView;
    OKBtn: TButton;
    CancelBtn: TButton;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure ListDblClick(Sender: TObject);
  private
  public
    function Execute(Client: TOpcSimpleClient;
      Options: TServerSearchOptions = []): Boolean;  {cf 1.15.3.1}

  end;

implementation

{$R *.DFM}

resourcestring
  SClientActive = 'Client cannot be active';
  SNoServers = 'There are no servers on this host';
  SPleaseSelect = 'Please select a server';

{ TSelectServerForm }

function TServerSelectDlg.Execute(Client: TOpcSimpleClient; Options: TServerSearchOptions): Boolean;
const
  NDATypes: array[TOpcDataAccessType] of String = ('DA1', 'DA2', 'DA3');
var
  SL: TOpcServerList;
  i: Integer;
  j: TOpcDataAccessType;
  dt: TOpcDataAccessTypes;
  DAStr: string;
begin
  try
    if Client.Active then
      raise EAbort.CreateRes(@SClientActive);
    List.Items.BeginUpdate;
    try
      GetOpcServers(Client.HostName, SL, Options);  {cf 1.15.3.1}
      if Length(SL) = 0 then
        raise EAbort.CreateRes(@SNoServers);
      List.Items.Clear;
      for i:= 0 to Length(SL) - 1 do
      with List.Items.Add do
      begin
//        Caption:= SL[i].ProgID;
        Caption:= SL[i].VerIndProgID;
        SubItems.Add(SL[i].UserType);
        Data := Pointer(i);
        dt:= SL[i].DaTypes;
        DAStr:= '';
        for j:= Low(TOpcDataAccessType) to High(TOpcDataAccessType) do
        if j in dt then
        begin
          if DAStr = '' then
            DAStr:= NDaTypes[j]
          else
            DAStr:= DAStr + ', ' + NDaTypes[j]
        end;
        SubItems.Add(DAStr);
//        SubItems.Add(SL[i].Vendor);
      end
    finally
      List.Items.EndUpdate
    end;
    Result:= ShowModal = mrOK;
    if Result and (List.SelCount = 1) then
      Client.ProgID:= SL[Integer(List.Selected.Data)].ProgID
  except
    on E:EAbort do
    begin
      Result:= false;
      MessageDlg(E.Message, mtError, [mbOK], 0)
    end
  end;
end;

procedure TServerSelectDlg.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOK) and (List.SelCount <> 1) then
  begin
    CanClose:= false;
    MessageDlg(SPleaseSelect, mtConfirmation, [mbOK], 0)
  end
end;

procedure TServerSelectDlg.ListDblClick(Sender: TObject);
begin
  OKBtn.Click;
end;

end.

