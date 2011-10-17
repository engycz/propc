{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
unit prOpcItemSelect;
{$I prOpcCompilerDirectives.inc}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Forms, Dialogs, Controls, StdCtrls,
  Buttons, ExtCtrls, ComCtrls, prOpcBrowser, prOpcDa, prOpcClient;

type
  TItemSelectDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Pages: TNotebook;
    SrcList: TListBox;
    SrcLabel: TLabel;
    IncludeBtn: TButton;
    IncAllBtn: TButton;
    ExcludeBtn: TButton;
    ExcAllBtn: TButton;
    DstList: TListBox;
    DstLabel: TLabel;
    PropertyView: TOpcPropertyView;
    PropertiesLabel: TLabel;
    BasicEditor: TMemo;
    procedure IncludeBtnClick(Sender: TObject);
    procedure ExcludeBtnClick(Sender: TObject);
    procedure IncAllBtnClick(Sender: TObject);
    procedure ExcAllBtnClick(Sender: TObject);
    procedure SrcListClick(Sender: TObject);
  private
    LastSelection: string;
    Client: TOpcSimpleClient;
    procedure UpdateSrcSelection;
    procedure MoveSelected(List: TCustomListBox; Items: TStrings);
    procedure SetItem(List: TListBox; Index: Integer);
    function GetFirstSelection(List: TCustomListBox): Integer;
    procedure SetButtons;
  public
    function Execute(aClient: TOpcSimpleClient; Group: TOpcGroup): Boolean;
  end;

var
  ItemSelectDlg: TItemSelectDlg;

implementation
resourcestring
  SHasOrphan = 'This item list contains %d entry which does not appear ' +
               'in the address space of the server. Do you want to keep it?';
  SHasOrphans = 'This item list contains %d entries which do not appear ' +
               'in the address space of the server. Do you want to keep them?';


{$R *.DFM}

procedure TItemSelectDlg.IncludeBtnClick(Sender: TObject);
var
  Index: Integer;
begin
  Index := GetFirstSelection(SrcList);
  MoveSelected(SrcList, DstList.Items);
  SetItem(SrcList, Index);
end;

procedure TItemSelectDlg.ExcludeBtnClick(Sender: TObject);
var
  Index: Integer;
begin
  Index := GetFirstSelection(DstList);
  MoveSelected(DstList, SrcList.Items);
  SetItem(DstList, Index);
end;

procedure TItemSelectDlg.IncAllBtnClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to SrcList.Items.Count - 1 do
    DstList.Items.AddObject(SrcList.Items[I],
      SrcList.Items.Objects[I]);
  SrcList.Items.Clear;
  SetItem(SrcList, 0)
end;

procedure TItemSelectDlg.ExcAllBtnClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to DstList.Items.Count - 1 do
    SrcList.Items.AddObject(DstList.Items[I], DstList.Items.Objects[I]);
  DstList.Items.Clear;
  SetItem(DstList, 0);
end;

procedure TItemSelectDlg.MoveSelected(List: TCustomListBox; Items: TStrings);
var
  I: Integer;
begin
  for I := List.Items.Count - 1 downto 0 do
    if List.Selected[I] then
    begin
      Items.AddObject(List.Items[I], List.Items.Objects[I]);
      List.Items.Delete(I);
    end;
end;

procedure TItemSelectDlg.SetButtons;
var
  SrcEmpty, DstEmpty: Boolean;
begin
  SrcEmpty := SrcList.Items.Count = 0;
  DstEmpty := DstList.Items.Count = 0;
  IncludeBtn.Enabled := not SrcEmpty;
  IncAllBtn.Enabled := not SrcEmpty;
  ExcludeBtn.Enabled := not DstEmpty;
  ExcAllBtn.Enabled := not DstEmpty;
end;

function TItemSelectDlg.GetFirstSelection(List: TCustomListBox): Integer;
begin
  for Result := 0 to List.Items.Count - 1 do
    if List.Selected[Result] then Exit;
  Result := LB_ERR;
end;

procedure TItemSelectDlg.SetItem(List: TListBox; Index: Integer);
var
  MaxIndex: Integer;
begin
  with List do
  begin
    SetFocus;
    MaxIndex := List.Items.Count - 1;
    if Index = LB_ERR then Index := 0
    else if Index > MaxIndex then Index := MaxIndex;
    Selected[Index] := True;
  end;
  SetButtons;
  UpdateSrcSelection
end;

procedure TItemSelectDlg.UpdateSrcSelection;
var
  NewSelection: string;
begin
  with SrcList do
  if SelCount = 1 then
    NewSelection:= Items[ItemIndex]
  else
    NewSelection:= '';
  if NewSelection <> LastSelection then
  begin
    LastSelection:= NewSelection;
    PropertyView.ShowProperties(Client, NewSelection)
  end
end;

function TItemSelectDlg.Execute(aClient: TOpcSimpleClient; Group: TOpcGroup): Boolean;
var
  Orphans: TStrings;
  i: Integer;
  Found: Integer;
  EntryString: String;
  Dest: TStrings;
begin
  Dest:= Group.Items;
  if aClient.SupportsBrowsing then
  begin
    Pages.PageIndex:= 1;
    Orphans:= TStringList.Create;
    try
      Client:= aClient;
      SrcList.Items.BeginUpdate;
      DstList.Items.BeginUpdate;
      try
        Client.GetAllItems(SrcList.Items);
        for i:= 0 to Dest.Count - 1 do
        begin
          Found:= SrcList.Items.IndexOf(Dest[i]);
          if Found <> -1 then
          begin
            SrcList.Items.Delete(Found);
            DstList.Items.Add(Dest[i])
          end else
          begin
            Orphans.Add(Dest[i])
          end
        end;
        if Orphans.Count > 0 then
        begin
          if Orphans.Count > 1 then
            EntryString:= SHasOrphans
          else
            EntryString:= SHasOrphan;
          case MessageDlg(Format(EntryString, [Orphans.Count]),
                      mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
            mrNo: Orphans.Clear;
            mrCancel: Abort;
          end
        end;
      finally
        SrcList.Items.EndUpdate;
        DstList.Items.EndUpdate
      end;
      SetButtons;
      Result:= ShowModal = mrOK;
      if Result then
      begin
        Dest.BeginUpdate;
        try
          Dest.Clear;
          Dest.AddStrings(DstList.Items);
          Dest.AddStrings(Orphans)
        finally
          Dest.EndUpdate
        end
      end
    finally
      Orphans.Free
    end
  end else  {basic mode}
  begin
    Pages.PageIndex:= 0;
    BasicEditor.Lines.Clear;
    BasicEditor.Lines.AddStrings(Dest);
    Result:= ShowModal = mrOK;
    if Result then
    with BasicEditor.Lines do
    begin
      Dest.BeginUpdate;
      try
        Dest.Clear;
        for i:= 0 to Count - 1 do
        if Strings[i] <> '' then
          Dest.Add(Strings[i])
      finally
        Dest.EndUpdate
      end
    end
  end
end;

procedure TItemSelectDlg.SrcListClick(Sender: TObject);
begin
  UpdateSrcSelection
end;

end.
