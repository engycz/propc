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
  TWizzyServer = class(TRttiItemServer)
  private
    function GetFormHeight: Integer;
    function GetFormWidth: Integer;
    procedure SetFormHeight(const Value: Integer);
    procedure SetFormWidth(const Value: Integer);
  protected
  public
  published
    {declare your Opc Items in here}
    property FormHeight: Integer read GetFormHeight write SetFormHeight;
    property FormWidth: Integer read GetFormWidth write SetFormWidth;
  end;

implementation
uses
  MainUnit,
  prOpcError;

{ TWizzyServer }

const
  ServerGuid: TGUID = '{F1B1B47F-BC2C-4E88-A579-F24125855B9B}';
  ServerVersion = 1;
  ServerDesc = 'Simple but Wizzy Server';
  ServerVendor = 'Wizzy Vendor';

{ TWizzyServer }

function TWizzyServer.GetFormHeight: Integer;
begin
  Result:= MainForm.Height
end;

function TWizzyServer.GetFormWidth: Integer;
begin
  Result:= MainForm.Width
end;

procedure TWizzyServer.SetFormHeight(const Value: Integer);
begin
  MainForm.Height:= Value
end;

procedure TWizzyServer.SetFormWidth(const Value: Integer);
begin
  MainForm.Width:= Value
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TWizzyServer.Create)
end.
