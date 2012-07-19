{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
{------------------------------------------------------------}
unit MainUnit;

interface
{prDemo12 - OpcServer as a service.
THIS DEMO WILL NOT WORK ON YOUR MACHINE - READ THE NOTES BELOW
--------------------------------------------------------------

There are a number of steps involved in setting up a an OPC server
as a service.

1. Select an account in which to run the service. To do this
set the ServiceStartName and Password properties of Demo12Service.
The Accountname should be in the form of DOMAIN\ACCOUNT. For a
'local machine' account use '.\ACCOUNT'. See the documentation for
the Win32 function 'CreateService' for more information. You will
need to set the 'Password' property to the correct password for the
account. If you do not do this, then your service will run in the
Local 'System' account. It may be possible to successfully connect
to COM objects instantiated in this account, but I do not know
how to do this (so there is no point in asking me - I have tried).

You must set ServiceStartName and Password to a valid user account
on your machine or your OPC server will not work.


2. Build the server.

3. Install the service by running prDemo12 /INSTALL. This will install
a new service called Demo12Service

4. Register the OPC server by running prDemo12 /REGSERVER. This will
register the server.

5. The 'standard' com registration procedure requires a bit of tinkering
to make the server run as a service, rather than as a normal application.
You must add a 'LocalService' value to the server's AppID key. The CLSID
for the TDemo12 object is C9161755-AE8A-4390-BD1E-F56EBFD29DE4. You will
find this declared ('ServerGuid') at the bottom of OpcServerUnit.pas.

The server's "AppId" key is thus :

}(*
HKEY_CLASSES_ROOT\AppID\{C9161755-AE8A-4390-BD1E-F56EBFD29DE4}
*){

Open RegEdit and locate this key (it will not be present until you
have registered your server) Add a StringValue called 'LocalService' and
set it to 'Demo12Service'. You can also do this by Merging the file
prDemo12LocalService.reg into your registry, which is included with this
example.

You should now be able to connect to your server. The 'ServiceStatus' item
shows the status of the service. Use

'net pause Demo12Service'
and
'net continue Demo12Service'

to change the status of the service. These state changes don't actually do
anything in this service. It is up to you to implement an appropriate
behaviour.

NOTE
----
This example is provided 'as-is' to help you get started with a Service
Application. I am not an expert in this type of Application. If you have
questions regarding the implementation or debugging of Service Applications,
rather than specific OPC questions please refer to other sources before you
ask me. I would suggest MSDN and/or usenet archives as good sources of
information - I assembled this example purely using these sources.

Also - I studiously avoid any issues to do with DCOM configuration and/or
security. These are general COM issues and nothing to do with OPC in
particular.
}


uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs;

type
  TDemo12Service = class(TService)
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
  private
    { Private declarations }
  public
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  Demo12Service: TDemo12Service;

implementation
uses
  prOpcServer;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  Demo12Service.Controller(CtrlCode);
end;

function TDemo12Service.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TDemo12Service.ServiceStop(Sender: TService;
  var Stopped: Boolean);
begin
  {cannot stop service while clients exist}
  if OpcItemServer.ClientCount > 0 then
    Stopped:= false
end;

end.
