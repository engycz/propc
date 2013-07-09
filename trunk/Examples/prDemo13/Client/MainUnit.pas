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
unit MainUnit;
{
  Client for prDemo13

  This client connects to prDemo13.exe which must be built and registered.

  The client has one group with one item. This item is defined in the server
  as an enumerated type which means that its 'Canonical Data Type' is
  VT_INTEGER. The client can, however, use the method
  TOpcGroup.ItemEnumeratedNames to retrieve a list of the names associated with
  the item. See ClientClient connect,
}


interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, prOpcClient, StdCtrls;

type
  TMainForm = class(TForm)
    ConnectButton: TButton;
    Client: TOpcSimpleClient;
    BorderStyleCB: TComboBox;
    procedure ConnectButtonClick(Sender: TObject);
    procedure ClientConnect(Sender: TObject);
    procedure ClientDisconnect(Sender: TObject);
    procedure BorderStyleCBSelect(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.ConnectButtonClick(Sender: TObject);
begin
  if Client.Active then
  begin
    Client.Active:= false;
    ConnectButton.Caption:= 'Connect'
  end else
  begin
    Client.Active:= true;
    ConnectButton.Caption:= 'Disconnect'
  end
end;

procedure TMainForm.ClientConnect(Sender: TObject);
begin
  {retrieve enumerated names from client}
  Client.Groups[0].ItemEnumeratedNames(0, BorderStyleCB.Items);
  BorderStyleCB.ItemIndex:= Client.Groups[0].ItemValue[0];
  BorderStyleCB.Enabled:= true
end;

procedure TMainForm.ClientDisconnect(Sender: TObject);
begin
  BorderStyleCB.Enabled:= false;
  BorderStyleCB.Items.Clear;
  BorderStyleCB.ItemIndex:= -1
end;

procedure TMainForm.BorderStyleCBSelect(Sender: TObject);
begin
  if BorderStyleCB.Enabled then
    Client.Groups[0].ItemValue[0]:= BorderStyleCB.ItemIndex
end;

end.
