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
unit prOpcWizMain;
{$I prOpcCompilerDirectives.inc}
interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, ToolsApi, ComCtrls;

type
  POpcServerData = ^TOpcServerData;
  THookType = (hkConnect, hkDisconnect, hkSetName, hkAddGroup,
        hkRemoveGroup, hkAddItem, hkRemoveItem);

  TOpcServerData = record
    Name, Desc, Vendor: String;
    Version: Integer;
    GUID: TGUID;
    SupportsSubscription: Boolean;
    MaxUpdateRate: Integer;  {0 for default}
    UseRtti, RecursiveRtti, HierarchicalBrowsing, ExtendedItemInfo: Boolean;
    Hooks: array[THookType] of Boolean;
    AddToSearchPath: Boolean;
  end;

  TOpcServerForm = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Label1: TLabel;
    Name: TEdit;
    Description: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Vendor: TEdit;
    Label4: TLabel;
    GenerateGUIDBox: TCheckBox;
    GUID: TEdit;
    SupportsSubscriptionBox: TCheckBox;
    VersionEdit: TEdit;
    VersionUD: TUpDown;
    Label5: TLabel;
    MaxUpdateEdit: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    DefaultUpdateBox: TCheckBox;
    AddToSearchPathBox: TCheckBox;
    HooksBox: TGroupBox;
    HookConnectBox: TCheckBox;
    HookDisconnectBox: TCheckBox;
    HookAddGroupBox: TCheckBox;
    HookRemoveGroupBox: TCheckBox;
    HookAddItemBox: TCheckBox;
    HookRemoveItemBox: TCheckBox;
    HookSetNameBox: TCheckBox;
    AllButton: TButton;
    NoneButton: TButton;
    OptionsBox: TGroupBox;
    UseRttiBox: TCheckBox;
    RecursiveRttiBox: TCheckBox;
    HierarchicalBox: TCheckBox;
    ExtendedInfoBox: TCheckBox;
    procedure GenerateGUIDBoxClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure DefaultUpdateBoxClick(Sender: TObject);
    procedure HooksButtonClick(Sender: TObject);
    procedure UseRttiBoxClick(Sender: TObject);
  private
  public
    Data: TOpcServerData;
    procedure GetServerData;
  end;

  TOpcServerWizard = class(TNotifierObject,
    IOTAWizard, IOTARepositoryWizard, IOTAFormWizard
{$IFDEF D9UP}
  ,IOTARepositoryWizard60, IOTARepositoryWizard80
{$ENDIF}
  )

  public
    function GetAuthor: string;
    function GetComment: string;
    function GetPage: string;
{$IFDEF D6UP}
    function GetGlyph: Cardinal;
{$ELSE}
    function GetGlyph: HICON;
{$ENDIF}
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    procedure Execute;
{$IFDEF D9UP}
  //IOTARepositoryWizard60 = interface(IOTARepositoryWizard)
  function GetDesigner: string;

  //IOTARepositoryWizard80 = interface(IOTARepositoryWizard60)
  function GetGalleryCategory: IOTAGalleryCategory;
  function GetPersonality: string;
{$ENDIF}


  end;

implementation
uses
  Dialogs, ComObj, ActiveX, TypInfo; {, DbugIntf;}

{$R *.DFM}

type
  TOpcFile = class(TInterfacedObject, IOTAFile)
    FSource: String;
    function GetSource: string;
    { Return the age of the file. -1 if new }
    function GetAge: TDateTime;
    constructor Create(const aSource: string);
  end;

  TModuleCreator = class(TInterfacedObject, IOTACreator, IOTAModuleCreator)
    FSource: string;
    function GetCreatorType: string;
    { Return False if this is a new module }
    function GetExisting: Boolean;
    { Return the File system IDString that this module uses for reading/writing }
    function GetFileSystem: string;
    { Return the Owning module, if one exists (for a project module, this would
      be a project; for a project this is a project group) }
    function GetOwner: IOTAModule;
    { Return true, if this item is to be marked as un-named.  This will force the
      save as dialog to appear the first time the user saves. }
    function GetUnnamed: Boolean;


    function GetAncestorName: string;
    { Return the implementation filename, or blank to have the IDE create a new
      unique one. (C++ .cpp file or Delphi unit) }
    function GetImplFileName: string;
    { Return the interface filename, or blank to have the IDE create a new
      unique one.  (C++ header) }
    function GetIntfFileName: string;
    { Return the form name }
    function GetFormName: string;
    { Return True to Make this module the main form of the given Owner/Project }
    function GetMainForm: Boolean;
    { Return True to show the form }
    function GetShowForm: Boolean;
    { Return True to show the source }
    function GetShowSource: Boolean;
    { Create and return the Form resource for this new module if applicable }
    function NewFormFile(const FormIdent, AncestorIdent: string): IOTAFile;
    { Create and return the Implementation source for this module. (C++ .cpp
      file or Delphi unit) }
    function NewImplSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile;
    { Create and return the Interface (C++ header) source for this module }
    function NewIntfSource(const ModuleIdent, FormIdent, AncestorIdent: string): IOTAFile;
    { Called when the new form/datamodule/custom module is created }
    procedure FormCreated(const FormEditor: IOTAFormEditor);

    constructor Create(const aSource: string);
  end;

  EOpcServerWizard = class(Exception);

const
  NHook: array[THookType] of string = (
    'OnClientConnect(Client: TClientInfo)',
    'OnClientDisconnect(Client: TClientInfo)',
    'OnClientSetName(Client: TClientInfo)',
    'OnAddGroup(Group: TGroupInfo)',
    'OnRemoveGroup(Group: TGroupInfo)',
    'OnAddItem(Item: TGroupItemInfo)',
    'OnRemoveItem(Item: TGroupItemInfo)');

  NHookHint: array[THookType] of string = (
    'connects',
    'disconnects',
    'calls IOpcCommon.SetClientName',
    'adds a group',
    'removes a group',
    'adds an item to a group',
    'removes an item from a group');

function GetUnitSource(const UnitName: string; WizData: TOpcServerData): String;
var
  s: TStrings;
  i: THookType;
begin
  s:= TStringList.Create;
  with s, WizData do
  try
    Add(Format('unit %s;', [UnitName]));
    Add('');
    Add('interface');
    Add('');
    Add('uses');
    if UseRtti then
      Add('  SysUtils, Classes, prOpcRttiServer, prOpcServer, prOpcTypes;')
    else
      Add('  SysUtils, Classes, prOpcServer, prOpcTypes;');
    Add('');
    Add(       'type');
    if UseRtti then
      Add(Format('  %s = class(TRttiItemServer)', [Name]))
    else
      Add(Format('  %s = class(TOpcItemServer)', [Name]));
    Add(       '  private');
    Add(       '  protected');
    if HierarchicalBrowsing then
      Add(     '    function Options: TServerOptions; override;');
    if UseRtti and not RecursiveRtti then
      Add(     '    function IsRecursive: Boolean; override;');
    if SupportsSubscription then
    begin
      Add(     '    function SubscribeToItem(ItemHandle: TItemHandle; UpdateEvent: TSubscriptionEvent): Boolean; override;');
      Add(     '    procedure UnsubscribeToItem(ItemHandle: TItemHandle); override;')
    end;
    for i:= Low(THookType) to High(THookType) do
    if Hooks[i] then
      Add(Format('    procedure %s; override;', [NHook[i]]));
    if MaxUpdateRate <> 0 then
      Add(     '    function MaxUpdateRate: Cardinal; override;');
    Add(       '  public');
    if UseRtti then
    begin
      Add(     '  published');
      Add(     '    {declare your Opc Items in here}');
    end else
    begin
      if ExtendedItemInfo then
      begin
        Add(   '    function GetExtendedItemInfo(const ItemID: String;');
        Add(   '      var AccessPath: string; var AccessRights: TAccessRights;');
        Add(   '      var EUInfo: IEUInfo; var ItemProperties: IItemProperties): Integer; override;');
      end else
      begin
        Add(   '    function GetItemInfo(const ItemID: String; var AccessPath: string;');
        Add(   '      var AccessRights: TAccessRights): Integer; override;')
      end;
      Add(     '    procedure ReleaseHandle(ItemHandle: TItemHandle); override;');
      Add(     '    procedure ListItemIDs(List: TItemIDList); override;');
      Add(     '    function GetItemValue(ItemHandle: TItemHandle;');
      Add(     '                            var Quality: Word): OleVariant; override;');
      Add(     '    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); override;');
    end;
    Add(       '  end;');
    Add('');
    Add('implementation');
    Add('uses');
    Add('  prOpcError;');
    Add('');
    Add(Format('{ %s }', [Name]));
    Add('');
    if HierarchicalBrowsing then
    begin
      Add(Format('function %s.Options: TServerOptions;', [Name]));
      Add(       'begin');
      Add(       '  Result:= inherited Options + [soHierarchicalBrowsing]');
      Add(       'end;');
      Add(       '')
    end;
    if UseRtti and not RecursiveRtti then
    begin
      Add(Format('function %s.IsRecursive: Boolean;', [Name]));
      Add(       'begin');
      Add(       '  Result:= false');
      Add(       'end;');
      Add(       '');
    end;
    if SupportsSubscription then
    begin
      Add(Format('function %s.SubscribeToItem(ItemHandle: TItemHandle; UpdateEvent: TSubscriptionEvent): Boolean;', [Name]));
      Add(       'begin');
      Add(       '  {Save UpdateEvent and call it whenever the Item referenced by ItemHandle changes}');
      Add(       '  Result:= false');
      Add(       '  {return true if subscription accepted}');
      Add(       'end;');
      Add(       '');
      Add(Format('procedure %s.UnsubscribeToItem(ItemHandle: TItemHandle);', [Name]));
      Add(       'begin');
      Add(       '  {Cancel the subscription started with a call to Subscribe to item}');
      Add(       'end;');
      Add(       '');
    end;
    for i:= Low(THookType) to High(THookType) do
    if Hooks[i] then
    begin
      Add(Format('procedure %s.%s;', [Name,NHook[i]]));
      Add(       'begin');
      Add(Format('  {Code here will execute whenever a client %s}', [NHookHint[i]]));
      Add(       'end;');
      Add(       '');
    end;
    if MaxUpdateRate > 0 then
    begin
      Add(Format('function %s.MaxUpdateRate: Cardinal;', [Name]));
      Add(       'begin');
      Add(Format('  Result:= %d' ,[MaxUpdateRate]));
      Add(       'end;');
      Add(       '');
    end;
    if not UseRtti then
    begin
      if ExtendedItemInfo then
      begin
        Add(Format('function %s.GetExtendedItemInfo(const ItemID: String;', [Name]));
        Add(       '  var AccessPath: string; var AccessRights: TAccessRights;');
        Add(       '  var EUInfo: IEUInfo; var ItemProperties: IItemProperties): Integer;');
      end else
      begin
        Add(Format('function %s.GetItemInfo(const ItemID: String; var AccessPath: string;', [Name]));
        Add(       '  var AccessRights: TAccessRights): Integer;')
      end;
      Add(       'begin');
      Add(       '  {Return a handle that will subsequently identify ItemID}');
      Add(       '  {raise exception of type EOpcError if Item ID not recognised}');
      Add(       '  raise EOpcError.Create(OPC_E_INVALIDITEMID)');
      Add(       'end;');
      Add(       '');
      Add(Format('procedure %s.ReleaseHandle(ItemHandle: TItemHandle);', [Name]));
      Add(       'begin');
      Add(       '  {Release the handle previously returned by GetItemInfo}');
      Add(       'end;');
      Add(       '');
      Add(Format('procedure %s.ListItemIds(List: TItemIDList);', [Name]));
      Add(       'begin');
      Add(       '  {Call List.AddItemId(ItemId, AccessRights, VarType) for each ItemId}');
      Add(       'end;');
      Add(       '');

      Add(Format('function %s.GetItemValue(ItemHandle: TItemHandle;', [Name]));
      Add(       '                           var Quality: Word): OleVariant;');
      Add(       'begin');
      Add(       '  {return the value of the item identified by ItemHandle}');
      Add(       'end;');
      Add(       '');

      Add(Format('procedure %s.SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant);', [Name]));
      Add(       'begin');
      Add(       '  {set the value of the item identified by ItemHandle}');
      Add(       'end;')
    end;
    Add(       '');
    Add('const');
    Add(Format('  ServerGuid: TGUID = ''%s'';', [GUIDToString(GUID)]));
    Add(Format('  ServerVersion = %d;', [Version]));
    Add(Format('  ServerDesc = ''%s'';', [Desc]));
    Add(Format('  ServerVendor = ''%s'';', [Vendor]));
    Add('');
    Add('initialization');
    Add(Format('  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, %s.Create)', [Name]));
    Add('end.');
    Result:= Text
  finally
    s.Free
  end
end;

procedure ShowSource( const ModuleServices: IOTAModuleServices);
var
  Module: IOTAModule;
  SourceEdit: IOTASourceEditor;
  Count: Integer;
  i: Integer;
begin
  Module:= ModuleServices.CurrentModule;
  Count:= Module.GetModuleFileCount;
  for i:= 0 to Count-1 do
  if Module.GetModuleFileEditor(i).QueryInterface(IOTASourceEditor, SourceEdit) = S_OK then
  begin
    SourceEdit.Show;
    break
  end
end;

function GetCurrentProject: IOTAProject;
var
  Services: IOTAModuleServices;
  Module: IOTAModule;
  Project: IOTAProject;
  ProjectGroup: IOTAProjectGroup;
  MultipleProjects: Boolean;
  I: Integer;
begin
  Result:= nil;
  MultipleProjects:= False;
  Services:= BorlandIDEServices as IOTAModuleServices;
  for I:= 0 to Services.ModuleCount - 1 do
  begin
    Module:= Services.Modules[I];
    if Module.QueryInterface(IOTAProjectGroup, ProjectGroup) = S_OK then
    begin
      Result:= ProjectGroup.ActiveProject;
      Exit
    end
    else if Module.QueryInterface(IOTAProject, Project) = S_OK then
    begin
      if Result = nil then
        Result := Project
      else
        MultipleProjects := True
    end
  end;
  if MultipleProjects then
    Result := nil
end;

type
  TStringListEx =
  class(TStringList)
  private
    function GetDelimitedText(Delimiter: Char): String;
    procedure SetDelimitedText(Delimiter: Char; const Value: String);
  public
    property DelimitedText[Delimiter: Char]: String read GetDelimitedText write SetDelimitedText;
  end;

function TStringListEx.GetDelimitedText(Delimiter: Char): string;
var
  i: Integer;
begin
  Result := '';
  for i:= 0 to Count - 1 do
    if i = 0 then
      Result:= Trim(Get(i))
    else
      Result:= Result + Delimiter + Trim(Get(i))
end;

procedure TStringListEx.SetDelimitedText(Delimiter: Char; const Value: string);
var
  P, P1: PChar;
  S: string;
begin
  BeginUpdate;
  try
    Clear;
    P:= PChar(Value);
    P1:= P;
    while P^ <> #0 do
    begin
      if P^ = Delimiter then
      begin
        SetString(S, P1, P - P1);
        Add(Trim(S));
        P1:= P + 1;
      end;
      Inc(P)
    end;
    SetString(S, P1, P - P1);
    Add(Trim(S));
  finally
    EndUpdate;
  end
end;

procedure TOpcServerWizard.Execute;

{The 'search path' in d5 seems to map to
  four older options ResDir, SrcDir, ObjDir, UnitDir
  Options:= GetCurrentProject.ProjectOptions;
  OV:= 'NEWPATH';
  Options.SetOptionValue('ResDir', OV);
  Options.SetOptionValue('SrcDir', OV);
  Options.SetOptionValue('ObjDir', OV);
  Options.SetOptionValue('UnitDir', OV)}

procedure AddToSearchPath;
var
  cp: IOTAProject;
  Options: IOTAProjectOptions;
  sl: TStringListEx;
  i: Integer;
  Found: Boolean;
  Path: String;
  NewPath: String;
const
  OptionNames: array[0..3] of String =
    ('ResDir', 'SrcDir', 'ObjDir', 'UnitDir');
  prOpcPath = '\prOpcKit';
  DelphiPath = '$(DELPHI)';

begin
  cp:= GetCurrentProject;
  if cp = nil then
  begin
    MessageDlg('Could not add to project search path as more than one project is active.',
       mtWarning, [mbOK], 0)
  end else
  begin
    {must be Delphi 5.01  - a bit complicated to check}
    Options:= cp.ProjectOptions;
    {use Unit Dir as base. This was the one I always used to set}
    {we need to check to see if prOpc is already on the search path.
     The user might have added the path manually and not necessarily using
     the ($DELPHI)\prOpc syntax, so if we find any path on the search path
     that ends in \prOpc we assume that the path is already there}
    sl:= TStringListEx.Create;
    try
      Path:= Trim(Options.GetOptionValue(OptionNames[3]));
      Found:= false;
      if Path <> '' then
      begin
        sl.DelimitedText[';']:= Path;
        for i:= 0 to sl.Count - 1 do
        begin
          Path:= Trim(sl[i]);
          if SameText(prOpcPath, Copy(Path, Length(Path) - Length(prOpcPath) + 1, MaxInt)) then
          begin
            Found:= true;
            break
          end
        end
      end;
      if not Found then
      begin
        sl.Add(DelphiPath + prOpcPath);
        NewPath:= sl.DelimitedText[';'];
        for i:= 0 to 3 do
          Options.SetOptionValue(OptionNames[i], NewPath)
      end;
    finally
      sl.Free
    end;
  end
end;

var
  WizForm: TOpcServerForm;
  WizData: TOpcServerData;
  MS: IOTAModuleServices;
  UnitName, ClassName, Filename: string;
begin
  Application.CreateForm(TOpcServerForm, WizForm);
  try
    if WizForm.ShowModal = mrOK then
    begin
      MS:= BorlandIDEServices as IOTAModuleServices;
      ShowSource(MS);
      WizData:= WizForm.Data;
      if WizData.AddToSearchPath then
        AddToSearchPath;
      MS.GetNewModuleAndClassName('', UnitName, ClassName, Filename);
      MS.CreateModule(TModuleCreator.Create(GetUnitSource(UnitName, WizData)))
    end
  finally
    WizForm.Free
  end
end;

function TOpcServerWizard.GetAuthor: string;
begin
  Result:= 'Production Robots Engineering Ltd'
end;

function TOpcServerWizard.GetComment: string;
begin
  Result:= 'Create an OPC Server module'
end;

{$IFDEF D6UP}
function TOpcServerWizard.GetGlyph: Cardinal;
{$ELSE}
function TOpcServerWizard.GetGlyph: HICON;
{$ENDIF}
begin
  Result:= LoadIcon(hInstance, 'OPCICON')
end;

function TOpcServerWizard.GetIDString: string;
begin
  Result:= 'Production Robots.Opc Server Wizard'
end;

function TOpcServerWizard.GetName: string;
begin
  Result:= 'Opc Server'
end;

function TOpcServerWizard.GetPage: string;
begin
{$IFDEF D9UP}
  Result:= 'prOpcKit'
{$ELSE}
  Result:= 'New'
{$ENDIF}
end;

function TOpcServerWizard.GetState: TWizardState;
begin
  Result:= []
end;

{$IFDEF D9UP}
function TOpcServerWizard.GetDesigner: string;
begin
  Result:= dVCL
end;

function TOpcServerWizard.GetGalleryCategory: IOTAGalleryCategory;
begin
  Result:= nil
end;

function TOpcServerWizard.GetPersonality: string;
begin
  Result:= sDelphiPersonality
end;
{$ENDIF}


procedure TOpcServerForm.GenerateGUIDBoxClick(Sender: TObject);
begin
  Guid.Enabled:= not GenerateGUIDBox.Checked
end;

procedure TOpcServerForm.GetServerData;
var
  i: Integer;

function CheckedEdit(Edit: TEdit): String;
begin
  if Edit.Text <> '' then
    Result:= Edit.Text
  else
    raise EOpcServerWizard.CreateFmt('You must enter a value for %s', [Edit.Name])
end;
begin
  Data.Name:= CheckedEdit(Name);
  if not IsValidIdent(Data.Name) then
    raise EOpcServerWizard.CreateFmt('%s is not a valid identifier', [Data.Name]);
  Data.Desc:= CheckedEdit(Description);
  Data.Vendor:= CheckedEdit(Vendor);
  Data.Version:= VersionUD.Position;
  if GenerateGUIDBox.Checked then
  begin
    CoCreateGUID(Data.GUID)
  end else
  try
    Data.GUID:= StringToGUID(Guid.Text)
  except
    on EOleSysError do
      raise EOpcServerWizard.CreateFmt('%s is not a valid GUID', [Guid.Text])
  end;
  with Data do
  begin
    if DefaultUpdateBox.Checked then
    begin
      MaxUpdateRate:= 0
    end else
    begin
      try
        MaxUpdateRate:= StrToInt(MaxUpdateEdit.Text)
      except
        on E: EConvertError do
          raise EOpcServerWizard.Create(E.Message)
      end;
      if MaxUpdateRate < 20 then
      begin
        if MessageDlg('An update rate of < 20ms is not practical. ' +
                        'Do you wish to use the '+#13+#10+'default value of 20ms?',
                         mtWarning, [mbOK, mbCancel], 0) = mrOK then
          MaxUpdateRate:= 0
        else
          Abort
      end
    end;
    Data.SupportsSubscription:= SupportsSubscriptionBox.Checked;
    Data.UseRtti:= UseRttiBox.Checked;
    Data.RecursiveRtti:= RecursiveRttiBox.Checked;
    Data.HierarchicalBrowsing:= HierarchicalBox.Checked;
    Data.ExtendedItemInfo:= ExtendedInfoBox.Checked;
    for i:= 0 to HooksBox.ControlCount - 1 do
    if HooksBox.Controls[i] is TCheckBox then
    with HooksBox.Controls[i] as TCheckBox do
      Data.Hooks[THookType(Tag)]:= Checked;
    Data.AddToSearchPath:= AddToSearchPathBox.Checked
  end
end;

(*
Creators are classes that you define and implement IOTACreator and one
of its derived interfaces. Delphi calls upon your creator objects to
create new files and proejcts.

IOTAFile
To supply the source code for a new file, you must define a class that
implements IOTAFile. Your creator creates an instance of your class and
returns that object from the appropriate method.

IOTACreator
Base interface for all creators. If you want to use a default creator,
e.g., for a default form or project, return an appropriate string as
CreatorType. To create a custom object, return an empty string.

IOTAModuleCreator
Create a unit, form, or other file. You can create a default unit, form,
or text file by returning sUnit, sForm, or sText as the CreatorType and
nil for the source file, form file, etc. Unlike the old-style module
creator, do not return empty strings as defaults for the file name, etc.
You must return the full information for the module. The form definition
is returned as an IOTAFile for a DFM file, including the DFM file signature.

IOTAProjectCreator
Create a new package, library, or application. You can create a default package,
library, or application by returning sPackage, sLibrary, or sApplication
as the CreatorType and nil for the new project source.
*)


{ TModuleCreator }

constructor TModuleCreator.Create(const aSource: string);
begin
  inherited Create;
  FSource:= aSource
end;

procedure TModuleCreator.FormCreated(const FormEditor: IOTAFormEditor);
begin

end;

function TModuleCreator.GetAncestorName: string;
begin
  Result:= ''
end;

function TModuleCreator.GetCreatorType: string;
begin
  Result:= ToolsApi.sUnit
end;

function TModuleCreator.GetExisting: Boolean;
begin
  Result:= false
end;

function TModuleCreator.GetFileSystem: string;
begin
  Result:= ''
end;

function TModuleCreator.GetFormName: string;
begin
  Result:= ''
end;

function TModuleCreator.GetImplFileName: string;
begin
  Result:= ''
end;

function TModuleCreator.GetIntfFileName: string;
begin
  Result:= ''
end;

function TModuleCreator.GetMainForm: Boolean;
begin
  Result:= false
end;

function TModuleCreator.GetOwner: IOTAModule;
var
  cp: IOTAProject;
begin
  cp:= GetCurrentProject;
  if Assigned(cp) then
    Result:= cp as IOTAModule
  else
    Result:= nil
end;

function TModuleCreator.GetShowForm: Boolean;
begin
  Result:= false
end;

function TModuleCreator.GetShowSource: Boolean;
begin
  Result:= true
end;

function TModuleCreator.GetUnnamed: Boolean;
begin
  Result:= true
end;

function TModuleCreator.NewFormFile(const FormIdent,
  AncestorIdent: string): IOTAFile;
begin
  Result:= nil
end;

function TModuleCreator.NewImplSource(const ModuleIdent, FormIdent,
  AncestorIdent: string): IOTAFile;
begin
  Result:= TOpcFile.Create(FSource)
end;

function TModuleCreator.NewIntfSource(const ModuleIdent, FormIdent,
  AncestorIdent: string): IOTAFile;
begin
  Result:= nil
end;

{ TOpcFile }

constructor TOpcFile.Create(const aSource: string);
begin
  inherited Create;
  FSource:= aSource
end;

function TOpcFile.GetAge: TDateTime;
begin
  Result:= -1
end;

function TOpcFile.GetSource: string;
begin
  Result:= FSource
end;

procedure TOpcServerForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose:= true;
  if ModalResult = mrOK then
  try
    GetServerData
  except
    on E: EOpcServerWizard do
    begin
      Application.ShowException(E);
      CanClose:= false
    end
  end
end;

procedure TOpcServerForm.DefaultUpdateBoxClick(Sender: TObject);
begin
  MaxUpdateEdit.Enabled:= not DefaultUpdateBox.Checked
end;

(*
procedure TOpcServerForm.GetButtonClick(Sender: TObject);
var
  OptNames: TOTAOptionNameArray;
  i: Integer;
  OV: String;
  Options: IOTAOptions;
begin
  Options:= GetCurrentProject.ProjectOptions;
  OptNames:= Options.GetOptionNames;
  for i:= Low(OptNames) to High(OptNames) do
  with OptNames[i] do
  begin
    if Kind in [tkString, tkLString] then
    begin
      OV:= Options.GetOptionValue(Name);
      SendDebug(Name + ' ' +
        GetEnumName(TypeInfo(TTypeKind), Ord(Kind)) + ' ' +
         OV)
    end
  end;
end;
*)

procedure TOpcServerForm.HooksButtonClick(Sender: TObject);
var
  SetVal: Boolean;
  i: Integer;
begin
  SetVal:= TButton(Sender).Tag = 1;
  for i:= 0 to HooksBox.ControlCount - 1 do
  if HooksBox.Controls[i] is TCheckBox then
    TCheckBox(HooksBox.Controls[i]).Checked:= SetVal
end;

procedure TOpcServerForm.UseRttiBoxClick(Sender: TObject);
begin
  RecursiveRttiBox.Enabled:= UseRttiBox.Checked;
  ExtendedInfoBox.Enabled:= not UseRttiBox.Checked
end;

end.
