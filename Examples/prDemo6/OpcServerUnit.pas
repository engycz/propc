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
  TDemo6 = class(TRttiItemServer)
  private
    function GetTickCount: Integer;
    function GetTimeOfDay: string;
  protected
  public
  published
    property TickCount: Integer read GetTickCount;
    property TimeOfDay: string read GetTimeOfDay;
  end;

implementation
uses
  prOpcError, Windows;


{ TDemo6 }

const
  ServerGuid: TGUID = '{E25AA401-121D-11D5-944C-00C0F023FA1C}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - In process demo';
  ServerVendor = 'Production Robots Eng. Ltd.';


{ TDemo6 }

function TDemo6.GetTickCount: Integer;
begin
  Result:= Windows.GetTickCount
end;

function TDemo6.GetTimeOfDay: string;
begin
  Result:= TimeToStr(Now)
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo6.Create)
end.
