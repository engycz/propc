{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit BrowseUnit;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, ComCtrls, prOpcBrowser, prOpcClient, prOpcTypes;

type
  TBrowseDlg = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ItemBrowser: TOpcBrowser;
    Splitter1: TSplitter;
    PropertyView: TOpcPropertyView;
    CloseButton: TButton;
  private
    { Private declarations }
  public
    procedure Execute(aClient: TOpcSimpleClient);
  end;

var
  BrowseDlg: TBrowseDlg;

implementation

{$R *.DFM}

{ TBrowseDlg }

procedure TBrowseDlg.Execute(aClient: TOpcSimpleClient);
begin
  ItemBrowser.OpcClient:= aClient;
  ShowModal;
  ItemBrowser.OpcClient:= nil
end;

end.
