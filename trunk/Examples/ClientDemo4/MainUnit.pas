{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000-2002  Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
{------------------------------------------------------------}
unit MainUnit;
{Very simple client demo}
{$IFDEF VER140}
  {$DEFINE D6UP}
{$ENDIF}
{$IFDEF VER150}
  {$DEFINE D6UP}
{$ENDIF}
interface

uses
{$IFDEF D6UP}
  Variants,
{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, prOpcVSC, StdCtrls, ExtCtrls, ComCtrls;

type
  TMainForm = class(TForm)
    StartButton: TButton;
    WriteButton: TButton;
    ReadButton: TButton;
    ReadResult: TLabel;
    WriteDataEdit: TEdit;
    WriteData: TUpDown;
    procedure StartButtonClick(Sender: TObject);
    procedure WriteButtonClick(Sender: TObject);
    procedure ReadButtonClick(Sender: TObject);
  private
  public
    VSC: Variant;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

{This project demonstrates the unit prOpcVSC

(VSC = VerySimpleClient)

prOpcVSC has only one exported function:

function OpcClient(
  const HostName, ProgId: string;
  UpdateRate: Integer): TVerySimpleClient;

Host name is the name of the server computer or '' for the local machine.
ProgId is the prog Id of the OPC server: for example 'prDemo5.TDemo5.1'
Update rate is the rate the server polls for data (ms) If update rate is
0 then the server does not poll and updates are not sent to the client. The
client can still do sync reads, however.

I reckon this is the simplest possible OPC client implementation.

To connect to a server:

VSC:= OpcClient('Machine', 'OpcServer.Serv.1', 0);

to set a tag 'TagName'

VSC.TagName:= 100;

to get a tag 'TagName'

int i;

i:= VSC.TagName;

TVerySimpleClient is really a variant, so to close the server set it to
Unassigned.

VSC:= Unassigned;   //closes connection to server

That's it. As might be expected the implementation is not wonderfully
efficient, but its not bad. If you specific an update rate of 0 then every read
goes to the server. If the update rate is not null, then the server sends
updates and the reading the client just reads a value stored on the client
side. This means that reads don't take a lot of time - thus the lack of a
mechanism to get callbacks is not such a big deal as polling a tag is
reasonably cheap.

Note however, that this client does a string lookup every time every time a tag
is dereferenced. The tag names are stored in a sorted string list and the lookup
is a binary search so this is pretty quick even if there lots of tag - but you
should use a local variable if you need to refer to a tag more than once in a
procedure.

Groups are (obviously) created 'behind the scenes'. If you call 'OpcClient'
twice on the same server, you don't get two connections, you just get two
groups on the same server.

To run this demo you will have to change the server to one that exists and
is registered on your machine.
}

procedure DrawTable(x, y, w, l: Integer);
begin
end;

procedure TestCode(Running: Boolean);
const
  ServerProgId = 'Snooker.Table.1';
var
  QuickChanging, SlowChanging: Variant;
  TableWidth, TableLength: Integer;

begin
  QuickChanging:= OpcClient('', ServerProgId, 40);  {40ms poll on server}
  SlowChanging:= OpcClient('', ServerProgId);      {synchronous reads only}
  TableWidth:= SlowChanging.TableWidth;
  TableLength:= SlowChanging.TableLength;
  while Running do
  begin
    DrawTable(QuickChanging.BallPositionX, QuickChanging.BallPositionY,
      TableWidth, TableLength);
    Sleep(100)
  end
end;


procedure TMainForm.StartButtonClick(Sender: TObject);
begin
  if VarIsEmpty(VSC) then
  begin
    VSC:= OpcClient('', 'prDemo5.TDemo5.1', 0);
    StartButton.Caption:= 'Stop';
    WriteButton.Enabled:= true;
    ReadButton.Enabled:= true
  end else
  begin
    VSC:= Unassigned;
    StartButton.Caption:= 'Start';
    WriteButton.Enabled:= false;
    ReadButton.Enabled:= false
  end
end;

procedure TMainForm.WriteButtonClick(Sender: TObject);
begin
  VSC.FormWidth:= WriteData.Position
end;

procedure TMainForm.ReadButtonClick(Sender: TObject);
begin
  ReadResult.Caption:= VSC.TimeOfDay
end;

end.
