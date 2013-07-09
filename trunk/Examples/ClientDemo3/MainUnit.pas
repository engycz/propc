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

interface
{release 1.15b. Fixed for delphi5}

uses
  Windows, Messages, SysUtils,
{$IFNDEF VER130}
  Variants,
{$ENDIF}
  Classes, Graphics, Controls, Forms,
  Dialogs, prOpcVSC, StdCtrls, ExtCtrls;

type
  TMainForm = class(TForm)
    StartButton: TButton;
    procedure StartButtonClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
  public
    VSC: Variant;
  end;

var
  MainForm: TMainForm;

implementation
{1. Make sure that prDemo5.exe is registered.
Run this program
Click Start
Resize the client form: Observe the server.
}


{$R *.dfm}

procedure TMainForm.StartButtonClick(Sender: TObject);
begin
  if VarIsEmpty(VSC) then
  begin
    VSC:= OpcClient('', 'prDemo5.TDemo5.1', 0);
    SetBounds(Left, Top, VSC.FormWidth, VSC.FormHeight);
    StartButton.Caption:= 'Stop'
  end else
  begin
    VSC:= Unassigned;
    StartButton.Caption:= 'Start'
  end
end;

procedure TMainForm.FormResize(Sender: TObject);
begin
  if not VarIsEmpty(VSC) then
  begin
    VSC.FormWidth:= Width;
    VSC.FormHeight:= Height
  end
end;

end.
