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
  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo1 = class(TRttiItemServer)
  private
    function GetFormHeight: Integer;
    function GetFormWidth: Integer;
    function GetTickCount: Integer;
    function GetTimeOfDay: String;
    procedure SetFormHeight(const Value: Integer);
    procedure SetFormWidth(const Value: Integer);
  protected
  public
  published
    {declare your Opc Items in here}
    property TickCount: Integer read GetTickCount;
    property TimeOfDay: String read GetTimeOfDay;
    property FormWidth: Integer read GetFormWidth write SetFormWidth;
    property FormHeight: Integer read GetFormHeight write SetFormHeight;
  end;

implementation
uses
  prOpcError, Windows, MainUnit;

{ TDemo1 }


const
  ServerGuid: TGUID = '{4113BAE2-130B-11D5-944D-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'Simple Demo using RTTI';
  ServerVendor = 'Production Robots Eng Ltd.';

{ TDemo1 }

function TDemo1.GetFormHeight: Integer;
begin
  Result:= MainForm.Height
end;

function TDemo1.GetFormWidth: Integer;
begin
  Result:= MainForm.Width
end;

function TDemo1.GetTickCount: Integer;
begin
  Result:= Windows.GetTickCount
end;

function TDemo1.GetTimeOfDay: String;
begin
  Result:= TimeToStr(Now)
end;

procedure TDemo1.SetFormHeight(const Value: Integer);
begin
  MainForm.Height:= Value
end;

procedure TDemo1.SetFormWidth(const Value: Integer);
begin
  MainForm.Width:= Value
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo1.Create)
end.

