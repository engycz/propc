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

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, prOpcVSC;

type
  TForm1 = class(TForm)
    WidthShrinkButton: TSpeedButton;
    WidthGrowButton: TSpeedButton;
    HeightShrinkButton: TSpeedButton;
    HeightGrowButton: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure WidthShrinkButtonClick(Sender: TObject);
    procedure HeightGrowButtonClick(Sender: TObject);
    procedure HeightShrinkButtonClick(Sender: TObject);
    procedure WidthGrowButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    VSC: Variant;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  VSC:= OpcClient('', 'Wizzy1.TWizzyServer.1')
end;

procedure TForm1.WidthGrowButtonClick(Sender: TObject);
begin
  VSC.FormWidth:= VSC.FormWidth + 20
end;

procedure TForm1.WidthShrinkButtonClick(Sender: TObject);
begin
  VSC.FormWidth:= VSC.FormWidth - 20
end;

procedure TForm1.HeightGrowButtonClick(Sender: TObject);
begin
  VSC.FormHeight:= VSC.FormHeight + 20
end;

procedure TForm1.HeightShrinkButtonClick(Sender: TObject);
begin
  VSC.FormHeight:= VSC.FormHeight - 20
end;


end.
