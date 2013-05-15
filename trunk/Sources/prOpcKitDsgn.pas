{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit prOpcKitDsgn;
{$I prOpcCompilerDirectives.inc}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ToolsApi, prOpcWizMain, prOpcClient, ComCtrls, StdCtrls, prOpcTypes, prOpcEnum,
  Buttons, prOpcBrowser, prOpcServerSelect, prOpcItemSelect;

procedure Register;

implementation
uses
{$IFDEF D6UP}
  DesignIntf,
  DesignEditors,
{$ELSE}
  DsgnIntf,
{$ENDIF}
  TypInfo;

type
  TProgIDProperty = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure Edit; override;
  end;

  TGroupItemsProperty = class(TPropertyEditor)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure Edit; override;
    function GetValue: String; override;
  end;

procedure Register;
begin
  RegisterComponents('prOPC', [TOpcSimpleClient, TOpcPropertyView, TOpcBrowser]);
  RegisterPropertyEditor(TypeInfo(string), TOpcSimpleClient, 'ProgID', TProgIDProperty);
  RegisterPropertyEditor(TypeInfo(TStrings), TOpcGroup, 'Items', TGroupItemsProperty);
  RegisterPackageWizard(TOpcServerWizard.Create)
end;

{ TProgIDProperty }

procedure TProgIDProperty.Edit;
var
  sf: TServerSelectDlg;
begin
  Application.CreateForm(TServerSelectDlg, sf);
  try
    if sf.Execute(TOpcSimpleClient(GetComponent(0))) then
      Modified
  finally
    sf.Free
  end
end;

function TProgIDProperty.GetAttributes: TPropertyAttributes;
begin
  Result:= inherited GetAttributes + [paDialog] - [paMultiSelect]
end;

{ TGroupItemsProperty }

procedure TGroupItemsProperty.Edit;
var
  Group: TOpcGroup;
  Client: TOpcSimpleClient;
  Form: TItemSelectDlg;
begin
  Group:= TOpcGroup(GetComponent(0));
  Client:= TOpcGroupCollection(Group.Collection).Client;
  Application.CreateForm(TItemSelectDlg, Form);
  try
    if Form.Execute(Client, Group) then
      Modified
  finally
    Form.Free
  end
end;

function TGroupItemsProperty.GetAttributes: TPropertyAttributes;
begin
  Result:= [paDialog, paReadOnly]
end;

function TGroupItemsProperty.GetValue: String;
begin
  Result:= 'Opc Items'
end;

end.

