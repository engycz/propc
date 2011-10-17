{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls;

type
  TMainForm = class(TForm)
    TreeView: TTreeView;
    DebugLog: TListBox;
    Splitter1: TSplitter;
    Panel1: TPanel;
    ShutdownClientsButton: TButton;
    ExitServerButton: TButton;
    procedure TreeViewDblClick(Sender: TObject);
    procedure ExitServerButtonClick(Sender: TObject);
    procedure ShutdownClientsButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation
uses
  prOpcServer, prOpcTypes;

{$R *.DFM}

procedure TMainForm.TreeViewDblClick(Sender: TObject);
const
  CR = #10#13;
  NBool: array[Boolean] of String = ('No', 'Yes');

  function TypeStr(VarType: Integer): String;
  begin
    case VarType of
      varEmpty: Result:= 'Empty';
      varNull: Result:= 'Null';
      varSmallint: Result:= 'Smallint';
      varInteger: Result:= 'Integer';
      varSingle: Result:= 'Single';
      varDouble: Result:= 'Double';
      varCurrency: Result:= 'Currency';
      varDate: Result:= 'Date';
      varOleStr: Result:= 'OleStr';
      varDispatch: Result:= 'Dispatch';
      varError: Result:= 'Error';
      varBoolean: Result:= 'Boolean';
      varVariant: Result:= 'Variant';
      varByte: Result:= 'Byte';
    else
      Result:= Format('Unknown type (%x)', [VarType])
    end
  end;

  function AccessStr(Rights: TAccessRights): string;
  const
    NAccessRight: array[TAccessRight] of String =
      ('Read', 'Write');
  var
    S: TStrings;
    j: TAccessRight;

  begin
    S:= TStringList.Create;
    try
      for j:= Low(TAccessRight) to High(TAccessRight) do
      if j in Rights then
        S.Add(NAccessRight[j]);
      Result:= S.CommaText
    finally
      S.Free
    end
  end;

  function GetClientInfo(ClientInfo: TClientInfo): String;
  var
    CN: String;
  begin
    with ClientInfo do
    begin
      if ClientName = '' then
        CN:= '(no name)'
      else
        CN:= ClientName;
      Result:= Format(
        'Name: %s' + CR +
        'Last update: %s' + CR +
        'State: %d',
        [CN,
         TimeToStr(FiletimeToDateTime(LastUpdateTime)),
         ServerState])
    end
  end;

  function GetGroupInfo(GroupInfo: TGroupInfo): String;
  begin
    with GroupInfo do
      Result:= Format(
        'Name: %s' + CR +
        'Public: %s' + CR +
        'Active: %s' + CR +
        'Update rate: %d' + CR +
        'ClientHandle: $%.8x' + CR +
        'Enabled: %s' + CR +
        'ItemCount: %d' + CR +
        'Connections' + CR +
        ' - DA2: %s' + CR +
        ' - DA1 (Data): %s' + CR +
        ' - DA1 (Datatime): %s' + CR +
        ' - DA1 (WriteComplete): %s',
       [Name,
        NBool[IsPublicGroup],
        NBool[Active],
        UpdateRate,
        hClientGroup,
        NBool[Enabled],
        ItemCount,
        NBool[DA2Connected],
        NBool[DA1Connected(da1Data)],
        NBool[DA1Connected(da1DataTime)],
        NBool[DA1Connected(da1WriteComplete)]]);
  end;

  function GetGroupItemInfo(GroupItem: TGroupItemInfo): String;
  begin
    with GroupItem do
      Result:= Format(
       'ItemID: %s' + CR +
       'ItemHandle: %.8x' + CR +
       'Client Handle: %.8x' + CR +
       'Active: %s' + CR +
       'Requested Type: %s' + CR +
       'Native Type: %s' + CR +
       'Last Update Time: %s' + CR +
       'Access Rights: %s',
       [ItemID,
        ItemHandle,
        hClient,
        NBool[Active],
        TypeStr(RequestedDataType),
        TypeStr(CanonicalDataType),
        TimeToStr(LastUpdateTime),
        AccessStr(AccessRights)])
  end;

var
  Info: String;
  Sel: TTreeNode;
  Obj: TObject;
begin
  Sel:= TreeView.Selected;
  if Assigned(Sel) and
     Assigned(Sel.Data) then
  begin
    Obj:= TObject(Sel.Data);
    if Obj is TClientInfo then
      Info:= GetClientInfo(TClientInfo(Obj))
    else
    if Obj is TGroupInfo then
      Info:= GetGroupInfo(TGroupInfo(Obj))
    else
    if Obj is TGroupItemInfo then
      Info:= GetGroupItemInfo(TGroupItemInfo(Obj))
    else
      raise Exception.Create('Unknown Info Object');
    MessageDlg(Info, mtInformation, [mbOK], 0)
  end
end;

procedure TMainForm.ExitServerButtonClick(Sender: TObject);
begin
  Close
end;

procedure TMainForm.ShutdownClientsButtonClick(Sender: TObject);
begin
  OpcItemServer.ShutdownRequest('prDemo5: server request shutdown')
end;

end.
