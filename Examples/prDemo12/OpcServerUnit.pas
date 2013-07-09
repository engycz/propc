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
  SysUtils, Classes, Windows, prOpcRttiServer, prOpcServer, prOpcTypes;

type
  TDemo12 = class(TRttiItemServer)
  private
    FReadWriteItem: Integer;
    function GetTimeOfDay: String;
    function GetTickCount: Cardinal;
    function GetServiceStatus: Integer;
  protected
  public
  published
    {declare your Opc Items in here}
    property TimeOfDay: String read GetTimeOfDay;
    property TickCount: Cardinal read GetTickCount;
    property ReadWriteItem: Integer read FReadWriteItem write FReadWriteItem;
    property ServiceStatus: Integer read GetServiceStatus;
  end;

implementation
uses
  MainUnit, prOpcError;

{ TDemo12 }

function TDemo12.GetServiceStatus: Integer;
begin
  if Assigned(Demo12Service) then
    Result:= Integer(Demo12Service.Status)
  else
    Result:= 0
end;

function TDemo12.GetTickCount: Cardinal;
begin
  Result:= Windows.GetTickCount
end;

function TDemo12.GetTimeOfDay: String;
begin
  Result:= DateToStr(Now)
end;


const
  ServerGuid: TGUID = '{C9161755-AE8A-4390-BD1E-F56EBFD29DE4}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit Demo12 - Service Application';
  ServerVendor = 'Production Robots Engineering Ltd';

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemo12.Create)
end.
