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
{prDemo15 - PercentageDeadband and Enumerated Types}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls;

type
  TMainForm = class(TForm)
    ExitServerButton: TButton;
    ShutdownClientsButton: TButton;
    TrackBar: TTrackBar;
    ListBox: TComboBox;
    procedure ExitServerButtonClick(Sender: TObject);
    procedure ShutdownClientsButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure ShowDeadband(LastVal: Integer; Deadband: Single);
  end;

var
  MainForm: TMainForm;

implementation
uses
  prOpcServer, prOpcTypes;

{$R *.DFM}

procedure TMainForm.ExitServerButtonClick(Sender: TObject);
begin
  Close
end;

procedure TMainForm.ShutdownClientsButtonClick(Sender: TObject);
begin
  OpcItemServer.ShutdownRequest('prDemo15: server request shutdown')
end;

procedure TMainForm.ShowDeadband(LastVal: Integer; Deadband: Single);
var
  IntDb: Integer;
begin
  IntDb:= Round(Deadband);
  TrackBar.SelStart:= LastVal - IntDb;
  TrackBar.SelEnd:= LastVal + IntDb
end;

end.

