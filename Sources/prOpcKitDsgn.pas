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

