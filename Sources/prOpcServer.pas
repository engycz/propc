{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{ mailto: prOpcKit@prel.co.uk                                }
{ http://www.prel.co.uk                                      }
{------------------------------------------------------------}
{History

Pre-release                         
24/02/00
added conditional define to force DA1 (for testing)

1.0 11/03/00
Added conditional define for evaluation copy.
* Evaluation copy supports <41 items  <4 Clients
* TServerImpl send OnDisconnect message after removing from ServerList
* DA1Registration now only occurs on RegServer and will create
   class key if necessary
* Added virtual method to allow changed ProgID
* Fixed bug in TOpcItemServer.ShutdownRequest
* Made ServerList a TThreadList.
* TClientInfo.ShutdownRequest returns a value to indicate if it worked.
* TOpcItemServer.ShutdownRequest returns how many clients it shutdown.
* Added TOpcItemServer.GetErrorString - I had completely forgotten about this!
* Two new Exception types EOpcError (for passing back to client via HRESULT)
  and EOpcServer (for raising an exception to be handled (or not!) within the
  server)
* Changed registration system to work a lot better. Also works with InProc
  servers. (use specialised factory)
* Fixed bug in strings enumerator - did not check pcelt for nil. Also changed
  declaration of ItemsEnumerator - pcelt was out not pointer. Could be passed
  as nil. (&&& marshalling issue?)

1.01 Maintenance release 03/03/01
ChangeForm  Description
1.01.1      Modified Behaviour of IOpcCommon.Get/SetLocaleID.
            ItemServers implementing GetErrorString will need to be updated,
            as LCID is now passed to this function.

1.01.2      OpcFoundation test program insists that Group.GetState returns
            bActive as 1 for TRUE. Delphi (quite correctly in my opinion) uses
            -1, which is the Ole standard for 16 and 32 bit booleans.
            I have changed this, but don't like it.

1.01.3      OpcFoundation test requires that PercentDeadband be saved and
            returned, even though use of it is optional. Field added
            to OpcGroupImpl. Also check check 0 <= Value <= 100.

1.01.4      AddGroup now returns UNSUPPORTEDRATE if requested rate is 0.

1.01.5      Implemented EnumConnectionPoints on OpcServer and OpcGroup. As
            if this interface serves any useful purpose...

1.01.6      GetItemID did not check that the value passed to it was a real
            ItemID.

1.01.7      GetItemAttributes now returns bActive as 1. This is a fudge to
            get around Opc Compliance test. I think the test is wrong, but
            I would say that, wouldn't I? See 1.01.2 above.

1.01.8      Do a basic test on RequestedDataType in AddItems. This just checks
            for impossible requests like arrays and so on. This is a requirement
            of OPC compliance test.

1.01.9      No longer use Gexperts for Group Lifetime Debug (use ODS instead)

1.01.10     SetState did not return OPC_S_UNSUPPORTEDRATE

1.01.11     IOpcGroup.SetName did not properly validate the new name

1.01.12     Blank (as opposed to null) names not allowed.

1.01.13     Items cloned with CloneGroup were created inactive. Changed.

Release 1.10
1.10.1      Array Support. GroupItemImpl.UpdateCache now uses
            prVarUtils.CompareVariant as the '=' operator fails on array-type
            variants

1.10.2      Array Support. Test for valid requested type now allows vtArray type
            Note that type coersion will always fail for array properties. I
            don't know what the foundation stance is on this. We now use the
            GlobalProcedure IsSupportedVarType added to prOpcTypes.pas

1.10.3      VariantChangeType will always fail for vtArray type variants so we now
            check for equality of Canonical and Requested data types before
            using this function in TGroupImpl.GetItemValue and GetCacheValue.
            Note that TGroupImpl.ItemSetDatatype will almost certainly fail for
            array items. I am not sure what the implications of this will be.

1.10.4      Remove TOpcItemServer.UseRecursiveRtti and add 'Options' this allows
            expansion in the future. Existing applications might need a tweak -
            also Wizard will need updating.

1.10.5      Preload Rtti information on creation of object

1.10.6      Changed all references to RTTI to Rtti. This is more readable, but
            should not affect code

1.10.7      Added support for TDateTime type as Rtti item.

Release 1.11
1.11.1      TGroupImpl now does a refresh on IConnectionPoint.Advise. The
            specification does not say whether this should be a cache or
            a device refresh and is vague about it anyway. The only reference
            is in section 4.5.6.3 on Refresh2
            Quote from specification:
            "The behavior of this function <Refresh2> is identical to what
            happens when Advise is called initially except that the
            OnDataChange Callback will include the transaction ID specified
            here. (The initial OnDataChange callback will contain a
            Transaction ID of 0)"

1.11.2      TGroupItemImpl.GetItemValue was not returning Timestamp.

1.11.3      GetItemProperties was not returning correctly typed variants
            (thanks to Frank v. Munchow-Pohl for spotting this)

Release 1.12
1.12.1      New option soAlwaysAllocateErrorArrays. Some clients insist that
            ppErrors is always allocated even if the call returns S_OK
            (Thus implying that all of ppErrors is S_OK). The spec is not
            clear on this.

1.12.2      D6 Compatible (uses Variants)

1.12.3      D6 is picky about assigning to uninitialised variants e.g. from the
            heap. Added function 'ZeroAllocation' to allocate initialised
            memory.

Release 1.13
1.13.1      Fixed bug in TGroupImpl.SyncIORead: hClient was not returned in
            OPCITEMSTATE record.

1.13.2      SetLocaleID now matches QueryAvailableLocaleIds

1.13.3      TServerImpl.RemoveGroup bad handle now returns E_INVALIDARG rather
            than E_OPC_INVALIDHANDLE. Roland's test program requires this

1.13.4      TServerImpl.AddGroup failed to remove group from group list if there
            was an exception - also a problem with CloneGroup. Fixed.

1.13.5      TServerImpl.BrowseOpcItemIds did not return S_FALSE when the
            Enumerator was empty. Also did not validate parameters.

1.13.6      TServerImpl.LookupItemIds returned S_OK instead of S_FALSE

1.13.7      TServerImpl.GetItemId should return E_INVALIDARG instead if
            E_OPC_INVALIDITEMID. Seems odd, but that's the spec.

1.13.8      TGroupImpl.CloneGroup did not return the requested interface.
            (just returned UNK)

1.13.9      Various methods did not check for bad count.

1.13.10     AsyncIO2 methods did not validate server handles. This
            is now implemented as a 'late' error - i.e reported in
            the callback. This is easier to implement and allowed by
            the spec

1.13.11     AsyncIO methods did not validate server handles. This is
            implemented in accordance with the spec - i.e any invalid
            handles abort callback.

1.13.12     When writing string type variables to a DataAccess 1.0 stream
            the strings were not checked for nil. This is a serious problem
            which will cause an AV in any DA1 server using string variables
            that might be null, i.e ''.

1.13.13     Added support for array types on DataAccess 1.0 streams

1.13.14     TGroupItemImpl.UpdateCache was updating cache on every tick.

1.13.15     Added New server option soIncludeNullInDA1ByteCount. When connected
            To DataAccess 1.0 clients, all updates are sent via a stream.
            String data is marshalled into the stream as a DWORD byte count,
            then the Data (as WideChars) then a Null terminator. Section
            4.6.4.6 of the Data Access specification states reasonably clearly
            state that the byte count should NOT include the null terminator.
            Previous versions of the toolkit did include the null. This was
            non-compliant but worked with most clients. This anomaly has been
            corrected, but this flag can be set to restore the previous
            behaviour if this breaks any existing servers that are required
            to work with non-compliant clients.

1.13.16     Rearranged TGroupItemImpl.SetItemValue a bit to try and make sure
            that no EVariantError exceptions escape from the message loop.
            This was happening with the terribly badly formed array variants
            generated by the Matrikon OPC Explorer on ASyncWrite.

Release 1.14
Some customers have asked for the ability to forcibly release interfaces
which have a non-zero reference count. This is most irregular behaviour and
not well supported by COM or Delphi COM support, but I have done my best to
implement it.

1.14.1      Catch OleSys Exceptions on calling Da2 OnDataChange
1.14.2      Errors calling event interfaces are routed through
            GroupCallbackError or ServerClientCallError
1.14.3      New type TClientCall to describe which client call has raised an
            exception.
1.14.4      Catch OleSys Exceptions on calling Da2 OnWriteComplete
1.14.5      Catch OleSys Exceptions on calling Da2 OnReadComplete
1.14.6      Catch OleSys Exceptions on calling Da1 DataChange
1.14.6      Catch OleSys Exceptions on calling Da1 WriteComplete
1.14.7      Some slight modifications to GetItemId
1.14.8      AddGroup was not initializing the default TimeBias correctly.
1.14.9      Bug in Foundation compliance test requires that GetEnable returns
            0 for FALSE and 1 for TRUE.

(pre-release to AH 14/11/01 here)

1.14.10     Added option soStrictServerValidation. Normally, server handle
            validation is done by catching an AV. This is very efficient,
            but it is not guaranteed to catch every possible invalid handle.
            With this flag set handle validation is much slower.
1.14.11     Added Define SLOWCALLBACKS to allow the compliance test to check
            the AsyncIO.Cancel2 interface.
1.14.12     It was possible to receive a Cancel callback while processing an
            Async task. In order to ensure that these calls are failed, we now
            delete Async task from Group TaskList before processing task.
1.14.13     Potentially very nasty heap corruption error if ASync task created
            with no active items - Async task was freed but not removed from
            Group's task list. May be related to 1.14.12
1.14.14     No longer uses SyncObjs. I forget why it was there in the first
            place.
1.14.15     Support for EU
1.14.16     Use run-time library version of TObjectList.
1.14.17     Support for percentage deadband.
1.14.18     Hierarchical Browsing
1.14.19     Remove global option flags and introduce public option properties
1.14.20     Eliminate all Rtti code from this unit
1.14.21     Flag for no browsing interface at all (soNoBrowsing)
1.14.22     Unhandled exceptions in the utility wndproc can kill the server
            in a really horrible way. This includes exceptions raised in
            notification routines. All exceptions are now handled in the
            WndProc. These exceptions are passed through 'UnhandledException'
            routine, which can be overridden if you care about these things.
1.14.23     Added state flags to GroupItemImpl to make sure that remove
            notification is not sent unless add notification has been sent.
1.14.24     Added define 'NoMasks' to allow version to be built for the
            'personal' edition of delphi.
1.14.25     Added 'LastUpdateValue' to ClientItemInfo
1.14.26     Changed all dynamic methods to virtual. There just is not enough
            of them to make it worthwhile.
1.14.27     Added a new debug hook - OnItemValueChange. This is called whenever
            an item's cache value is updated. I wanted this for the deadband
            example and I thought there might be other cases where it could be
            useful. I have not added it to the wizard (too obscure).
1.14.28     Added virtual GetTimestamp to allow servers to use their own
            timesource.

Release 1.14a
1.14.29     Added 'Data' to TNamespaceNode
1.14.30     TItemIdList.AddItemId now returns TNamespaceNode.
1.14.31     GOpcItemServer now initialised in constructor of TOpcItemServer

Release 1.15a
1.15.2      Fixed bug in ChangeBrowsePosition. Null value of szString should
            Browse to root.

release 1.15b Maintenance release to AH.
1.15.3      Fixed memory leak in TRefreshTask

release 1.15.5  (115e)
1.15.5.1    added function to release resources associated with client without
            attempting to release callback interfaces.

1.15.5.2.   Correct bug whereby ActiveList was being freed before processing of
            Asynchronous Refresh2 tasks - rather nasty, this one.


release 1.15.6 (115g)
1.15.6.1    TItemIdList now ignores duplicates to allow InvalidateNamespace with
            'keep existing'

1.15.6.2    Fixed AV in InvalidateNamespace

}

unit prOpcServer;
{$I prOpcCompilerDirectives.inc}
interface
uses
  Windows, Messages, SysUtils, Classes, ActiveX, ComObj,
  Contnrs, {cf 1.14.16}
  prOpcComn, prOpcDa, prOpcError, prOpcTypes, prOpcUtils;

type
  TSubscriptionEvent = procedure (const Value: OleVariant; Quality: Word; TimeStamp: TFileTime) of object;

  TItemHandle = Integer;   {i.e foreign to this unit}

  {cf 1.14.18}
  TItemIdList = class;
  TNamespaceNode = class
  private
    FParent: TItemIdList;
    FName: string;
  protected
  public
    Data: Pointer;  {1.14.29}
    constructor Create(aParent: TItemIdList; const aName: String);
    function Path: string;
    function Child(i: Integer): TNamespaceNode; virtual;
    function ChildCount: Integer; virtual;
    property Name: String read FName;
    property Parent: TItemIdList read FParent;
  end;

  TNamespaceItem = class(TNamespaceNode)
  private
    FAccessRights: TAccessRights;
    FVarType: Integer;
  public
    constructor Create(aParent: TItemIdList; const aName: String;
      aAccessRights: TAccessRights; aVarType: Integer);
    property AccessRights: TAccessRights read FAccessRights;
    property VarType: Integer read FVarType;
  end;

  {This class is passed to the ListItemIDs function to enable browsing
  of available items by OPC clients  = NamespaceBranch. Name of class
  is awkward but has to be maintained for backwards compatibility}
  TItemIDList = class(TNamespaceNode)
  private
    FChildren: TStringList;
    procedure Clear;
    procedure AddChild(Child: TNamespaceNode);
  protected
  public
    constructor Create(aParent: TItemIdList; const aName: String);
    destructor Destroy; override;
    function Find(const Path: string): TNamespaceNode;
    function AddItemID(const ItemID: String;
                        AccessRights: TAccessRights;
                        VarType: Integer): TNamespaceNode; {cf 1.14.30}
    function NewBranch(const aName: String): TItemIDList;
    function Child(i: Integer): TNamespaceNode; override;
    function ChildCount: Integer; override;
  end;

  TClientInfo = class;
  TGroupInfo = class;
  TGroupItemInfo = class;

  TServerOption = (
    soAlwaysAllocateErrorArrays,
    soIncludeNullInDA1ByteCount,
    soStrictHandleValidation,    {cf 1.14.10}
    soNoBrowsing,                {cf 1.14.21}
    soHierarchicalBrowsing);     {cf 1.14.18}

  {cf 1.14.3}
  TGroupCallback = (
    ccDa2DataChange,
    ccDa2ReadComplete,
    ccDa2WriteComplete,
    ccDa2CancelCompete, {not used in this unit}
    ccDa1DataChange,
    ccDa1WriteComplete,
    ccDa1ReadComplete);

  TClientCallback = (
    ccShutdownRequest);

  TServerOptions = set of TServerOption;

  TPrivateItemList = class
    function LockList: TStringList; virtual; abstract;
    procedure UnlockList; virtual; abstract;
  end;

  {This is the main item server which defines available OPC items and
  their values. This is an abstract class which is implemented to
  create a working server}
{$M+}
  TOpcItemServer = class
  private
    FServerList: TThreadList;
    FItemList: TPrivateItemList;
    FVendorInfo: String;
    FStartupTime: TFiletime;
    FRootNode: TItemIdList; {cf 1.14.18}
    FIgnoreDuplicatesInListItemIds: Boolean;
    procedure InitBrowsing(var CurrentNode: TItemIdList);
    function GetServerOption(Option: TServerOption): Boolean; {cf 1.14.19}
  protected
    procedure UnhandledException(E: TObject); virtual; {cf 1.14.22}
    function PathDelimiter: Char; virtual;
    function PropertyDelimiter: Char; virtual;

    function GetErrorString(Code: HResult; LCID: DWORD): String; virtual; {cf 1.01.1}
    {raises E_INVALID_ARG if code not recognised}

    function Options: TServerOptions; virtual;

    function MaxUpdateRate: Cardinal; virtual;
    {If you want to limit the rate at which your data is polled then overwrite this function.
     Value in milliseconds. The default value is 20 milliseconds}

    function SubscribeToItem(ItemHandle: TItemHandle; UpdateEvent: TSubscriptionEvent): Boolean; virtual;
    {if you don't want the opc server to poll your data, then overwrite this function and
    1) return TRUE to indicate that callbacks are active
    2) call the function UpdateEvent whenever the value or quality of the data item changes

    If you do not implement subscription, then your data will be polled at the rate requested by
    its OPC clients}

    procedure UnsubscribeToItem(ItemHandle: TItemHandle); virtual;
    {If you implement SubscribeToItem you MUST implement this as well. This will be called by the server
     when it does not want callbacks any more. If you cannot implement this then don't implement Subscribe
     to Item and rely on polling}

    function FilterFunction(const FilterMask: String;
                            FilterAccessRights: TAccessRights;
                            FilterVarType: TVarType;
                            const ItemID: String;
                            ItemAccessRights: TAccessRights;
                            ItemVarType: TVarType): Boolean; virtual;

    function GetServerID( Version: Integer): string; virtual;
    {This is the bit of the ProgID after the <Exe name>. <GetServerVersionIndependentID>.Version by
     default}

    function GetServerVersionIndependentID: string; virtual;
    {This is the bit of the ProgID after the Exe name. ClassName by
     default}

    function GetItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights): Integer; virtual;
    {GetItemInfo must Return an 'ItemHandle' value which will be used to refer
    to the item in subsequent calls to subscribe, unsubscribe, get item value
    etc.. etc..

    If there is any problem (e.g unrecognised ItemID) then GetItemInfo must raise
    an exception of type EOpcError and return an appropriate code

    Return Code	          Description
    -----------           -----------
    OPC_E_INVALIDITEMID   The ItemID is not syntactically valid
    OPC_E_UNKNOWNITEMID   The ItemID is not in the server address space
    E_FAIL                The function was unsuccessful. (presumably for some other reason)

    In order to specify access rights edit the AccessRights parameter. This parameter is
    set to OPC_READABLE or OPC_WRITABLE by default

    This function will not be called twice for the same ItemID even if there are lots of
    OPCItems subscribed to it.

    Ignore AccessPath. It is not supported but included to allow for support to
    be added in the future without breaking clients}

    function GetExtendedItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights;
                        var EUInfo: IEUInfo;
                        var ItemProperties: IItemProperties): Integer; virtual;
    {GetExtendedItemInfo may be implemented instead of GetItemInfo for servers
    that implement EUInfo or ItemProperties}

    procedure ReleaseHandle(ItemHandle: TItemHandle); virtual;
    {This is called when the Server no longer requires the handle}

    procedure ListItemIDs(List: TItemIDList); virtual;
    {browse address space. Add all valid ItemID's to List.
    Use List.AddItemID to add the items}

    function GetItemVQT(ItemHandle: TItemHandle; var Quality: Word;
      var Timestamp: TFileTime): OleVariant; virtual;

    function GetItemValue(ItemHandle: TItemHandle;
                          var Quality: Word): OleVariant; virtual; 
    {Note: the result is returned as a Variant. The Canonical data type of an item should not change
     during its lifetime. This is not checked}

    procedure SetItemVQT(ItemHandle: TItemHandle; const ValueVQT: OPCITEMVQT); virtual;
    procedure SetItemValue(ItemHandle: TItemHandle; const Value: OleVariant); virtual; abstract;
    procedure SetItemQuality(ItemHandle: TItemHandle; const Quality: Word); virtual;
    procedure SetItemTimestamp(ItemHandle: TItemHandle; const Timestamp: TFileTime); virtual;

    procedure OnClientConnect(aServer: TClientInfo); virtual;
    procedure OnClientDisconnect(aServer: TClientInfo); virtual;
    procedure OnClientSetName(aServer: TClientInfo); virtual;
    procedure OnAddGroup(Group: TGroupInfo); virtual;
    procedure OnRemoveGroup(Group: TGroupInfo); virtual;
    procedure OnAddItem(Item: TGroupItemInfo); virtual;
    procedure OnRemoveItem(Item: TGroupItemInfo); virtual;
    procedure OnItemValueChange(Item: TGroupItemInfo); virtual;    {cf 1.14.27}

    procedure GroupCallbackError(Exception: EOleSysError; Group: TGroupInfo; Call: TGroupCallback); virtual; {cf 1.14.2}
    procedure ClientCallbackError(Exception: EOleSysError; Server: TClientInfo; Call: TClientCallback); virtual; {cf 1.14.2}

    {property support
    procedure GetPropertyInfoFromItemId(const ItemId: string; var ParentId: string; var Pid: Integer); virtual;
    function GetItemIdFromPropertyInfo(const ParentId: string; Pid: Integer): string; virtual;
    function GetPropertyInfo(const ItemId: string): TItemProperties; virtual; }

    property ItemList: TPrivateItemList read FItemList;

    {GetTimestamp. By default this uses the PC Clock but servers might want to
    use an external timesource}
    procedure GetTimestamp(var Timestamp: TFileTime); virtual; {cf 1.14.28}
  public
    constructor Create;
    destructor Destroy; override;
    procedure InvalidateNamespace(KeepExisting: Boolean); {cf 1.14.18}
    function RootNode: TItemIdList;  {cf 1.14.18}
    function ShutdownRequest(const Reason: String; ClientNo: Integer = -1): Integer;
    function ClientCount: Integer;
    function ClientInfo(i: Integer): TClientInfo;
    procedure ParseItemId(const ItemId: string; Path: TStrings); overload; {cf 1.14.18}
    procedure ParseItemId(const ItemId: string; Path: TStrings; var PropName: string); overload; {cf 1.14.18}
    property StartupTime: TFiletime read FStartupTime;

    {cf 1.14.19}
    property AlwaysAllocateErrorArrays: Boolean index soAlwaysAllocateErrorArrays read GetServerOption; {cf. 1.12.1}
    property IncludeNullInDA1ByteCount: Boolean index soIncludeNullInDA1ByteCount read GetServerOption; {cf. 1.13.15}
    property StrictHandleValidation: Boolean index soStrictHandleValidation read GetServerOption; {cf 14.1.10}
    property NoBrowsing: Boolean index soNoBrowsing read GetServerOption;     {cf 1.14.21}
    property HierarchicalBrowsing: Boolean index soHierarchicalBrowsing read GetServerOption;     {cf 1.14.18}
  published
  end;
{$M-}

  {TClientInfo, TGroupInfo and TGroupItemInfo are created and
   maintained by OpcClients. They are declared as public
   classes in this unit to provide diagnostic information.
   These objects are not owned by the server}

  TGroupItemInfo = class
  private
    FActive: Boolean;
    FRequestedDataType: TVarType;
    FhClient: OPCHANDLE;
    function GetItemID: String;
    function GetItemHandle: TItemHandle;
    function GetCanonicalDataType: TVarType;
    function GetAccessRights: TAccessRights;
    function GetServerItem: TObject; virtual; abstract;
  public
    Data: Pointer; {reserved for use by ItemServer descendents. Not used by prOpcKit}
    property Active: Boolean read FActive;
    property RequestedDataType: TVarType read FRequestedDataType;
    property hClient: OPCHANDLE read FhClient;
    property ItemID: String read GetItemID;
    property ItemHandle: TItemHandle read GetItemHandle;
    property CanonicalDataType: TVarType read GetCanonicalDataType;
    property AccessRights: TAccessRights read GetAccessRights;
    function Group: TGroupInfo; virtual; abstract;
    function LastUpdateTime: TDateTime; virtual; abstract;
    function LastUpdateValue: OleVariant; virtual; abstract;
    function ItemEUInfo: IEUInfo;
    function ItemProperties: IItemProperties;
  end;

  TDa1Format = (da1Data, da1DataTime, da1WriteComplete);

  TGroupInfo = class
  private
    FName: String;
    FClientInfo: TClientInfo;
    FIsPublicGroup: Boolean; {always false in this unit}
    FActive: Boolean;
    FUpdateRate: DWORD;
    FhClientGroup: OPCHANDLE;
    FLCID: TLCID;
    FDataChangeEnable: Boolean;
    FDeleted: Boolean;
    FDataCallback: IOPCDataCallback;
    FPercentDeadband: Single;
    FDa1Advise: array[TDa1Format] of IAdviseSink;
  public
    Data: Pointer; {reserved for use by ItemServer descendents. Not used by prOpcKit}
    property Name: String read FName;
    property IsPublicGroup: Boolean read FIsPublicGroup;
    property Active: Boolean read FActive;
    property UpdateRate: DWORD read FUpdateRate;
    property PercentDeadband: Single read FPercentDeadband;  {cf 1.01.3, 1.14.17}
    property hClientGroup: OPCHANDLE read FhClientGroup;
    property LCID: TLCID read FLCID;
    property Enabled: Boolean read FDataChangeEnable;
    property Deleted: Boolean read FDeleted;
    property ClientInfo: TClientInfo read FClientInfo;
    function DA2Connected: Boolean;
    function DA1Connected(Da1Format: TDa1Format): Boolean;
    function ItemCount: Integer; virtual; abstract;
    function Item(i: Integer): TGroupItemInfo; virtual; abstract;
  end;

  {this is NOT a connection to a remote server - it is a model
  of the remote client}
  TClientInfo = class(TComObject)
  private
    FClientName: WideString;
    FLCID: TLCID; {not used at present &&&}
    FOPCShutdown: IOPCShutdown;
    FLastUpdateTime: TFiletime;
    FServerState: OPCSERVERSTATE;
  public
    Data: Pointer; {reserved for use by ItemServer descendents. Not used by prOpcKit}
    property ClientName: WideString read FClientName;
    property LCID: TLCID read FLCID;
    property LastUpdateTime: TFiletime read FLastUpdateTime;
    property ServerState: OPCSERVERSTATE read FServerState;
    function ShutdownRequestAvail: Boolean;
    function ShutdownRequest(const Reason: String): Boolean;
    function GroupCount: Integer; virtual; abstract;
    function Group(i: Integer): TGroupInfo; virtual; abstract;
    procedure ForceDisconnect; virtual; abstract;   {cf 1.15.5.1}
  end;

  EOpcServer = class(Exception);  {to be handled by server}

  EOpcError = class(Exception)   {to be returned to client}
  private
    FErrorCode: HRESULT;
    function GetErrorCode: HRESULT;
  public
    constructor Create(aErrorCode: HRESULT);
    property ErrorCode: HRESULT read GetErrorCode;
  end;

const
  AllAccess = [iaRead, iaWrite];
  TimestampNotSet : TFileTime = ();

procedure RegisterOPCServer(const ServerGUID: TGUID;
                               ServerVersion: Integer;
                            const ServerDesc: String;
                          const ServerVendor: String;
                              aOpcItemServer: TOpcItemServer);

{Note ProgID is  ExeName.OpcItemsClassName.ServerVersion
  eg. "MyOpcServerProgram.TMyOpcItems.1" }

function OpcItemServer: TOpcItemServer;

implementation

{Casual Notes
 ------------
There is no support for server side blobs. (what are they actually for?)
}

{Defines:

GLD: "Group Lifetime Debug" sends messages to debugger to confirm group creation
and destroy. Used to check that group reference counting is working properly

Evaluation: For evaluation version. Imposes limits on number of clients/items

SlowCallbacks: Causes Async returns to be somewhat delayed. The OpcFoundation
compliance test cannot do all its tests if async callbacks are too fast
(notably cancel). There is no need to use this define in a production server

ForceDA1: Creates a DA1 server with no DA2 interfaces
}

uses
{$IFDEF D6UP}
  Variants,
{$ENDIF}
  TypInfo, ComServ,
{$IFNDEF NoMasks}
  Masks,
{$ENDIF}
  prOpcVarUtils;
  {Do not include AxCtrls: it is a huge unit to include, and most of the
   apparently useful stuff in there won't work with non-automation interfaces}

type
  TItemResult = record
    ItemHandle: TItemHandle;
    CanonicalDataType: TVarType;
    AccessRights: DWORD;
  end;

  TServerImpl = class;
  TGroupImpl = class;
  TGroupItemImpl = class;
  TServerItemList = class;

  {TServerItem
   -----------
  Owned by OpcItemServer's sorted stringlist (FItemList).
  There is only one of these for each unique ItemID. All access
  to OpcItemServer goes though here. There may be more than one
  TGroupItem which has a reference to this ServerItem - a pointer
  to each TGroupItem holding a reference to this ServerItem is
  held in RefList - This acts as Reference Count - when RefList.Count
  drops to 0 the TServerItem is destroyed (and removed from
  OpcItemServer.FItemList) A server item is thus always associated
  with at least one GroupItem - hence the GroupItem in the constructor

  Note that the groupitems in the reflist may all come from different
  clients - if the server is apartment threaded this means that there
  will be some synchronisation issues with TServerItem.

  ** issues regarding RefList **
  I have vacillated between making RefList a TThreadList. There
  is no doubt that access to this list must be serialised, but
  it seems unreasonable for every Server Item to have its own
  critical section.

  AddRef and ReleaseRef may be protected by a common critical section,
  because they are not called that much and it doesn't really matter.

  ItemCallback is a bit more complicated. By going through a single
  critical section, all calls through FDataCallback are serialised,
  even though they might be going to different clients. This seems
  unreasonable however in practice, the program will probably have
  only one thread that makes these callbacks anyway, so the practical
  effect will be small. On balance, I think I shall use one common
  Critical Section to protect all the reference lists. This is
  TServerItemList.FRefListCritSect. I might change this later

  Polling Callbacks etc..
  -----------------------
  (recall that some server items will be polled, and some not)

  Why is cache value stored in the GroupItem and not the ServerItem?
  There is a good reason for this. The cache value should be
  the last value sent to the client. If there are two GroupItems
  connected to the same ServerItem they both need to keep independent
  records of the last value sent to the client so they can correctly
  determine if an update is necessary.

  Callbacks Enabled:
  If callbacks are enabled on the item then when the callback reaches
  the GroupItem the function UpdateCache is called. If this returns true
  then the cache value has changed and the group item is marked for update
  by setting GroupItem.CacheUpdated:= true. The next time the group is
  polled, the item is added to the update list and the cache values are
  sent to the client via FDataCallback.

  Callbacks not enabled:
  When the Group is polled, if callbacks are not enabled then the
  ServerItem.GetItemValue is called to get a new value. This is then
  passed to UpdateCache to check to see if the item needs to be added
  to the update list

  Note that if callbacks are not enabled and the same server item exists
  more than once in a group then ServerItem.GetItemValue may be called
  more than once per tick. This is inefficient but probably quite
  unusual.

  See TGroupItem.Tick

  Question: Should the Group tick occur if Group is not connected?
  Answer: Probably.

  There are server items which are Not associated with group items.
  These are generated in response to casual enquiries from interfaces
  IOPCBrowseServerAddressSpace and IOPCItemProperties
  }

  TServerItemRef = class
    ItemID: String;
    ItemHandle: DWORD;
    AccessRights: DWORD;
    CanonicalDataType: TVarType;
    EUInfo: IEUInfo;
    ItemProperties: IItemProperties;
    CacheValue: OleVariant;
    CacheTimestamp: TFiletime;
    CacheQuality: Word;
    AnalogType: Boolean;  {true if type is suitable for use of deadband.
                           arrays are excluded for now cf 1.14.17}
    procedure GetMaxAgeVQT(MaxAge: Cardinal; ActualTimestamp: TFileTime; var Result: OleVariant;
      var Quality: Word; var Timestamp: TFileTime);
    procedure GetCacheVQT(var Result: OleVariant; var Quality: Word; var Timestamp: TFileTime);
    procedure GetItemVQT(var Result: OleVariant; var Quality: Word; var Timestamp: TFileTime);
    procedure SetItemVQT(const ValueVQT: OPCITEMVQT);
    procedure SetItemValue(const Value: OleVariant);
    procedure AssignTo(var Result: OPCITEMATTRIBUTES);
    constructor Create(const aItemID: String);
    destructor Destroy; override;
    class function GetItem(const aItemID: String): TServerItemRef;
    {if exists in Global server item list then return item in list
    otherwise create}
    procedure ReleaseNonGroupReference; virtual; {free}
  end;

  TServerItem = class(TServerItemRef)
    Owner: TServerItemList;
    RefList: TList;
    Subscribed: Boolean;
    procedure AddRef(GroupItem: TGroupItemImpl);
    procedure ReleaseRef(GroupItem: TGroupItemImpl);
    procedure ItemCallback(const Value: OleVariant; Quality: Word; TimeStamp: TFileTime);
    function LockRefList: TList;
    procedure UnlockRefList;
    constructor Create(const aItemID: String;
                       aOwner: TServerItemList;
                       aGroupItem: TGroupItemImpl);
    destructor Destroy; override;
    procedure ReleaseNonGroupReference; override; {don't free}
  end;

  {Different Clients need to refer to this list to get references to
  server items}

  TServerItemList = class(TPrivateItemList)
    FCritSect, FRefListCritSect: TRtlCriticalSection;
    FList: TStringList;
    constructor Create;
    destructor Destroy; override;
    function LockList: TStringList; override;
    procedure UnlockList; override;
    procedure DeleteItem(const ItemID: String);
    function AddServerItem(GroupItem: TGroupItemImpl; const ItemID: String): TServerItem;
  end;

  TConnectionPoint = class(TObject, IUnknown, IConnectionPoint)
    FContainer: Pointer;
    FIID: TGUID;
    FSink: IUnknown;
    FOnConnect: TConnectEvent;
    { IUnknown }
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
    { IConnectionPoint }
    function GetConnectionInterface(out iid: TIID): HResult; stdcall;
    function GetConnectionPointContainer(
      out cpc: IConnectionPointContainer): HResult; stdcall;
    function Advise(const unkSink: IUnknown; out dwCookie: Longint): HResult; stdcall;
    function Unadvise(dwCookie: Longint): HResult; stdcall;
    function EnumConnections(out enumconn: IEnumConnections): HResult; stdcall;
    constructor Create(const Container: IUnknown;
      const IID: TGUID; OnConnect: TConnectEvent);
    destructor Destroy; override;
  end;

  TGroupItemState = (gisNotified);               {cf 1.14.23}
  TGroupItemStates = set of TGroupItemState;     {cf 1.14.23}

  {Server Handle passed back to client is a pointer to TGroupItem}
  TGroupItemImpl = class(TGroupItemInfo)
  private
    Owner: TGroupImpl;
    ServerItem: TServerItem;
    CacheUpdated: Boolean; {by subscription}
    CacheValue: OleVariant;
    CacheTimestamp: TFiletime;
    CacheQuality: Word;
    {EUType: OPCEUTYPE;  }    {not at present &&&}
    {EUInfo: OleVariant; }
    GroupItemState: TGroupItemStates; {cf 1.14.23}
    PercentDeadband: Single;
    PercentDeadbandSet: Boolean;
    function GetServerItem: TObject; override;
    procedure SetActive(Value: Boolean);
    constructor CreateClone(aOwner: TGroupImpl; Source: TGroupItemImpl);
    procedure AssignTo(var Result: OPCITEMATTRIBUTES);
    procedure ItemCallback(const Value: OleVariant; Quality: Word; Timestamp: TFiletime);
    procedure GetItemValue(var Result: OleVariant; var Quality: Word; var Timestamp: TFiletime);
    procedure GetCacheValue(var Result: OleVariant; var Quality: Word; var Timestamp: TFiletime);
    procedure GetMaxAgeValue(MaxAge: Cardinal; ActualTimestamp: TFileTime; var Result: OleVariant;
      var Quality: Word; var Timestamp: TFiletime);
    procedure InvalidateCache; {load cache from device and mark for update. This should be done
                                when the item is activated - either at group or item level}
    function UpdateCache(const Value: OleVariant; Quality: Word; Timestamp: TFileTime): Boolean;
    function Tick: Boolean;  {returns true if item needs sending to client}
  public
    function Group: TGroupInfo; override;
    function LastUpdateTime: TDateTime; override;
    function LastUpdateValue: OleVariant; override;
    constructor Create(aOwner: TGroupImpl; const ItemDef: OPCITEMDEF;
                       var ItemResult: OPCITEMRESULT);
    destructor Destroy; override;
  end;

  {this should be an ordered list with a binary search
   in order to make hServer validation more efficient &&&
   see comments elsewhere re hServer Validation}
  TGroupItemList = class(TObjectList)
    function GetGroupItem(i: Integer): TGroupItemImpl;
    property GroupItem[i: Integer]: TGroupItemImpl read GetGroupItem; default;
  end;

  {this interface is implemented by TGroupImpl so that it can be
  passed to TItemEnumerator, which is passed out to clients.
  We don't want any object pointers in TItemEnumerator as TGroupImpl
  might be released while ItemEnumerators exist}
  IGroupItemList = interface(IUnknown)
    function List: TGroupItemList;
  end;

  {note: this is a slightly non standard enumerator - the standard
  enumerator interfaces used in this unit are IEnumString and
  IEnumUnknown. In those cases, the output of Next is an array of
  pointers allocated by the client (in the case of strings, the
  individual strings returned are allocated by the server, although
  I am not sure this is correct &&&). In the case of IEnumOPCItemAttributes
  ppItemArray is a pointer to a pointer and the entire structure is
  allocated by the server... odd...}
  TItemEnumerator = class(TInterfacedObject, IEnumOPCItemAttributes)
    FIndex: Integer;
    GroupItemList: IGroupItemList;
    constructor Create(const aGroupItemList: IGroupItemList);
    function Next(celt: ULONG; out ppItemArray: POPCITEMATTRIBUTESARRAY;
      pceltFetched: PULONG): HResult; stdcall;
    function Skip(celt: Cardinal): HResult; stdcall;
    function Reset:HResult; stdcall;
    function Clone(out ppEnumItemAttributes: IEnumOPCItemAttributes): HResult; stdcall;
  end;

  TValueArray = array of OleVariant;

  {Async Task Processing.
   The task is allocated by the group, which adds the task to the
   pending tasks list. It is then posted to the opc window message queue
   as CM_ASYNCTASK (wParam is a pointer to the task). The window proc simply
   frees the Task. This carries out any processing (in the destructor) including
   removing from the group's list. Tasks that are cancelled simply have their deleted
   flag set so that when they are handled (destroyed) no action is taken}
  TAsyncTask = class
    Deleted: Boolean;
    BlockProcess: Boolean;
    TransactionID: DWORD;
    Group: TGroupImpl;
    constructor Create(aGroup: TGroupImpl; aTransactionID: DWORD);
    procedure DoProcess;  {cf 1.15.5.2}
    procedure Process; virtual; abstract;
    destructor Destroy; override;
  end;

  TRefreshTask = class(TAsyncTask)
    Source: OPCDATASOURCE;
    ActiveList: TList;
    constructor Create(aGroup: TGroupImpl; aTransactionID: DWORD;
      aSource: OPCDATASOURCE);
    procedure Process; override;
    destructor Destroy; override;  {cf 1.15.3}
  end;

  TRefreshMaxAgeTask = class(TAsyncTask)
    MaxAge: DWORD;
    ActualTimestamp: TFileTime;
    ActiveList: TList;
    constructor Create(aGroup: TGroupImpl; aTransactionID: DWORD;
      aMaxAge: DWORD);
    procedure Process; override;
    destructor Destroy; override;
  end;

  TRefreshTask1 = class(TRefreshTask)
    Format: TDa1Format;
    constructor Create(aGroup: TGroupImpl; aSource: OPCDATASOURCE;
      aFormat: TDa1Format);
    procedure Process; override;
  end;

  TAsyncIOTask = class(TAsyncTask)
    GroupItem: array of TGroupItemImpl;
    GroupItemValid: array of Boolean;
    Count: Integer;
    ValidCount: Integer;
    constructor Create(aGroup: TGroupImpl; aTransactionID, aCount: DWORD;
      phServer: POPCHANDLEARRAY; ppErrors: PResultList; var Result: HResult);
    function ProcessItems(ppErrors: PResultList; var Data): HRESULT;
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); virtual; abstract;
  end;

  TAsyncReadTask = class(TAsyncIOTask)
    procedure Process; override;
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); override;
  end;

  TAsyncReadMaxAgeTask = class(TAsyncIOTask)
    MaxAge : array of DWORD;
    ActualTimestamp: TFileTime;
    constructor Create(aGroup: TGroupImpl; aTransactionID, aCount: DWORD;
      phServer: POPCHANDLEARRAY; pdwMaxAge: PDWORDARRAY; ppErrors: PResultList;
      var Result: HResult);
    procedure Process; override;
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); override;
  end;

  TDataChangeStream = class(TMemoryStream)
    procedure WriteVariant(const Value: OleVariant);
    function Format: TDa1Format; virtual; abstract;
    function ItemHeaderSize: Integer; virtual; abstract;
    function GroupHeader: POPCGROUPHEADER;
    constructor Create(aCount: Integer);
    function ItemHeader2(i: Integer): POPCITEMHEADER2;
    procedure SetTimestamp(i: Integer; const aTimestamp: TFiletime); virtual;
  end;

  {with Timestamp}
  TDataChangeStream1 = class(TDataChangeStream)
    function Format: TDa1Format; override;
    function ItemHeaderSize: Integer; override;
    function ItemHeader1(i: Integer): POPCITEMHEADER1;
    procedure SetTimestamp(i: Integer; const aTimestamp: TFiletime); override;
  end;

  {without timestamp}
  TDataChangeStream2 = class(TDataChangeStream)
    function Format: TDa1Format; override;
    function ItemHeaderSize: Integer; override;
  end;

  TAsyncReadTask1 = class(TAsyncIOTask)
    Format: TDa1Format;
    Source: OPCDATASOURCE;
    function CreateStream: TDataChangeStream;
    procedure Process; override;
    constructor Create(aGroup: TGroupImpl; aCount: DWORD;
      phServer: POPCHANDLEARRAY; aSource: OPCDATASOURCE; aFormat: TDa1Format;
      ppErrors: PResultList; var Result: HResult);
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); override;
  end;

  TAsyncWriteTask = class(TAsyncIOTask)
    Values: TValueArray;
    constructor Create(aGroup: TGroupImpl; aTransactionID, aCount: DWORD;
      phServer: POPCHANDLEARRAY; pValues: POleVariantArray; ppErrors: PResultList;
      var Result: HResult);
    procedure Process; override;
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); override;
  end;

  TAsyncWriteVQTTask = class(TAsyncIOTask)
    Values: array of OPCITEMVQT;
    constructor Create(aGroup: TGroupImpl; aTransactionID, aCount: DWORD;
      phServer: POPCHANDLEARRAY; pValues: POPCITEMVQTARRAY; ppErrors: PResultList;
      var Result: HResult);
    procedure Process; override;
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); override;
  end;

  TAsyncWriteTask1 = class(TAsyncWriteTask) {Data Access 1.0 calls}
    constructor Create(aGroup: TGroupImpl; aCount: DWORD;
      phServer: POPCHANDLEARRAY; pValues: POleVariantArray; ppErrors: PResultList;
      var Result: HResult);
    procedure Process; override;
    procedure ProcessItem(Item: TGroupItemImpl; IndexSrc, IndexDst: Integer;
      var Data; var Status: HRESULT); override;
  end;

  TItemAction = procedure(Item: TGroupItemImpl; Index: Integer; Data: Pointer) of object;

  {Group Properties (from AddGroup documentation)
  szName
  ------
  Name of the group. The name must be unique among the other groups created by this client.
  If no name is provided (szName is pointer to a NUL string) the server will generate a
  unique name. The server generated name will also be unique relative to any existing public
  groups.

  bActive
  -------
  FALSE  if the Group is to be created as inactive.
  TRUE if the Group is to be created as active.

  dwRequestedUpdateRate
  ---------------------
  Client Specifies the fastest rate at which data changes may be sent to OnDataChange for items
  in this group. This also indicates the desired accuracy of Cached Data. This is intended only
  to control the behavior of the interface. How the server deals with the update rate and how
  often it actually polls the hardware internally is an implementation detail.  Passing 0 indicates
  the server should use the fastest practical rate.  The rate is specified in milliseconds.

  hClientGroup
  ------------
  Client provided handle for this group.
  [refer to description of data types, parameters, and structures for more information about
  this parameter]

  pTimeBias
  ---------
  Pointer to Long containing the initial TimeBias (in minutes) for the Group.  Pass a NULL Pointer
  if you wish the group to use the default system TimeBias. See discussion of TimeBias in
  General Properties Section See Comments below.

  pPercentDeadband
  ----------------
  The percent change in an item value that will cause a subscription callback for that value to a
  client. This parameter only applies to items in the group that have dwEUType of Analog.
  [See discussion of Percent Deadband in General Properties Section]. A NULL pointer is
  equivalent to 0.0.

  dwLCID
  ------
  The language to be used by the server when returning values (including EU enumeration's) as text
  for operations on this group.  This could also include such things as alarm or status conditions
  or digital contact states.

  phServerGroup
  -------------
  Place to store the unique server generated handle to the newly created group. The client will use
  the server provided handle for many of the subsequent functions that the client requests the server
  to perform on the group.

  pRevisedUpdateRate
  ------------------
  The server returns the value it will actually use for the UpdateRate which may differ from the
  RequestedUpdateRate.

  Note that this may also be slower than the rate at which the server is internally obtaining the
  data and updating the cache.
  In general the server should 'round up' the requested rate to the next available supported rate.
  The rate is specified in milliseconds.  Server returns HRESULT of OPC_S_UNSUPPORTEDRATE
  when it returns a value in revisedUpdateRate that is different than RequestedUpdateRate}

  TGroupList = class(TStringList)
  public
    constructor Create;
    destructor Destroy; override;
    procedure UpdateNames;
    procedure GroupByName(const Name: String; const riid: TIID; out ppUnk: IUnknown);
    function Group(i: Integer): TGroupImpl;
    function GetUniqueGroupName: String;
    function FindGroup(const Name: String): Boolean;
    procedure AddGroup(const Name: String; aGroup: TGroupImpl);
  end;

  TGroupState = (gsTimerRunning, gsNotified);
  TGroupStates = set of TGroupState;

  TGroupKeepAlive = class;

  TGroupImpl = class(TGroupInfo, IUnknown, IOPCItemMgt, IOPCGroupStateMgt,
                    IOPCSyncIO,
                    {$IFNDEF FORCEDA1}
                    IConnectionPointContainer, IOPCAsyncIO2,
                    {$ENDIF}
                    IGroupItemList,
                    {DA1 support} IDataObject, IOpcAsyncIO,
                    {DA3 support} IOPCGroupStateMgt2, IOPCSyncIO2, IOPCAsyncIO3, IOPCItemDeadbandMgt)
  private
    ReferenceCount: Integer;
    TimeBias: Longint;
    FItemList: TGroupItemList;
    FTaskList: TList;
    FConnectionPoint: TConnectionPoint;
    FGroupState: TGroupStates;
    FGroupKeepAlive: TGroupKeepAlive;
    procedure SetActive(Value: Boolean);
    procedure SetUpdateRate(Value: DWORD);
    function AddOrValidate( DoAdd, BlobUpdate: Boolean;
      dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
      var ppResults: POPCITEMRESULTARRAY;
      var ppErrors: PResultList): HResult;
    procedure ItemSetActiveState(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSetClientHandle(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSetDatatype(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemRemove(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSyncIORead(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSyncIOWrite(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSyncIO2Read(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSyncIO2Write(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemSetDeadband(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemGetDeadband(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    procedure ItemClearDeadband(Item: TGroupItemImpl; Index: Integer; Data: Pointer);
    function IterateItems(dwCount: Integer; phServer: POPCHANDLEARRAY; Data: Pointer;
       Action: TItemAction; var ppErrors: PResultList): HRESULT;
{    procedure ItemCallback( Item: TGroupItemImpl; const Value: OleVariant; Quality: Word);}
    procedure SinkConnect(const Sink: IUnknown; Connecting: Boolean);
    function DeleteTask(Task: TASyncTask): Boolean;
    procedure DoRefresh2( Source: OPCDATASOURCE; TransactionID: DWORD; ActiveList: TList);
    procedure DoRefreshMaxAge(MaxAge: DWORD; ActualTimestamp: TFileTime; TransactionID: DWORD; ActiveList: TList);
    procedure DoRefresh1(Source: OPCDATASOURCE; TransactionID: DWORD;
      Format: TDa1Format; ActiveList: TList);
    function InScope(dwScope: OPCENUMSCOPE): Boolean;
    procedure CheckDeleted;
    procedure CheckConnected;
    procedure CheckDa1Connected(dwConnection: Integer);
    procedure Tick;
    procedure ValidateServerHandle( hServer: OPCHANDLE);
    procedure SetPercentDeadband( Value: Single);
    function AllocErrorArray(dwCount: DWORD): PResultList; {cf. 1.12.1}
    function ASyncBadHandle(dwCount: DWORD; phServer: POPCHANDLEARRAY;
                 var ppErrors: PResultList): Boolean;

    {IUnknown}   {we do our own reference counting as RemoveGroup is complicated...}
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    { IOPCItemMgt }
    function AddItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY; out ppAddResults: POPCITEMRESULTARRAY;
                     out ppErrors: PResultList): HResult; stdcall;
    function ValidateItems(dwCount:DWORD; pItemArray:POPCITEMDEFARRAY; bBlobUpdate: BOOL; out ppValidationResults: POPCITEMRESULTARRAY;
                          out ppErrors: PResultList): HResult; stdcall;
    function RemoveItems(dwCount: DWORD; phServer: POPCHANDLEARRAY; out ppErrors: PResultList): HResult; stdcall;
    function SetActiveState(dwCount: DWORD; phServer: POPCHANDLEARRAY; bActive: BOOL; out ppErrors: PResultList): HResult; stdcall;
    function SetClientHandles(dwCount: DWORD; phServer: POPCHANDLEARRAY; phClient: POPCHANDLEARRAY;
                             out ppErrors: PResultList): HResult; stdcall;
    function SetDatatypes(dwCount: DWORD; phServer: POPCHANDLEARRAY; pRequestedDatatypes: PVarTypeList;
                         out ppErrors: PResultList): HResult; stdcall;
    function CreateEnumerator(const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;

    { IOPCGroupStateMgt }
    function GetState(out pUpdateRate: DWORD; out pActive: BOOL; out ppName:POleStr;
                     out pTimeBias:Longint; out pPercentDeadband:Single; out pLCID:TLCID;
                     out phClientGroup:OPCHANDLE; out phServerGroup: OPCHANDLE): HResult; stdcall;
    function SetState(pRequestedUpdateRate:PDWORD; out pRevisedUpdateRate: DWORD; pActive:PBOOL;
                     pTimeBias: PLongint; pPercentDeadband: PSingle; pLCID:PLCID;
                     phClientGroup: POPCHANDLE): HResult; stdcall;
    function SetName(szName: POleStr): HResult; stdcall;
    function CloneGroup(szName: POleStr; const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;

    { IOPCGroupStateMgt2 }
    function SetKeepAlive(dwKeepAliveTime: DWORD; out pdwRevisedKeepAliveTime: DWORD): HResult; stdcall;
    function GetKeepAlive(out pdwKeepAliveTime: DWORD): HResult; stdcall;

    { OPCSyncIO }
    function IOPCSyncIO.Read = SyncIORead;
    function SyncIORead(dwSource: OPCDATASOURCE; dwCount: DWORD; phServer: POPCHANDLEARRAY;
                 out ppItemValues: POPCITEMSTATEARRAY; out ppErrors: PResultList): HResult; stdcall;
    function IOPCSyncIO.Write = SyncIOWrite;
    function SyncIOWrite(dwCount:DWORD; phServer:POPCHANDLEARRAY; pItemValues:POleVariantArray;
                  out ppErrors:PResultList):HResult; stdcall;

    { IOPCSyncIO2 }
    function IOPCSyncIO2.Read = SyncIORead;
    function IOPCSyncIO2.Write = SyncIOWrite;
    function IOPCSyncIO2.ReadMaxAge = SyncIO2ReadMaxAge;
    function SyncIO2ReadMaxAge(dwCount: DWORD; phServer: POPCHANDLEARRAY; pdwMaxAge: PDWORDARRAY;
                              out ppvValues: POleVariantArray; out ppwQualities: PWordArray;
                              out ppftTimeStamps: PFileTimeArray; out ppErrors: PResultList): HResult; stdcall;
    function IOPCSyncIO2.WriteVQT = SyncIO2WriteVQT;
    function SyncIO2WriteVQT(dwCount: DWORD; phServer: POPCHANDLEARRAY; pItemVQT: POPCITEMVQTARRAY;
                            out ppErrors: PResultList): HResult; stdcall;

    { IOPCAsyncIO2 }
    {$IFNDEF FORCEDA1}
    function IOPCAsyncIO2.Read = AsyncIO2Read;
    {$ENDIF}
    function AsyncIO2Read(dwCount: DWORD; phServer: POPCHANDLEARRAY; dwTransactionID:DWORD;
                 out pdwCancelID: DWORD; out ppErrors: PResultList): HResult; stdcall;
    {$IFNDEF FORCEDA1}
    function IOPCAsyncIO2.Write = AsyncIO2Write;
    {$ENDIF}
    function AsyncIO2Write(dwCount: DWORD; phServer: POPCHANDLEARRAY; pItemValues: POleVariantArray;
                  dwTransactionID:DWORD; out pdwCancelID: DWORD; out ppErrors: PResultList): HResult; stdcall;
    function Refresh2(dwSource: OPCDATASOURCE; dwTransactionID:DWORD;
                     out pdwCancelID: DWORD): HResult; stdcall;
    function Cancel2(dwCancelID: DWORD): HResult; stdcall;
    function SetEnable(bEnable: BOOL): HResult; stdcall;
    function GetEnable(out pbEnable: BOOL): HResult; stdcall;

    { IOPCAsyncIO3 }
    function IOPCAsyncIO3.ReadMaxAge = AsyncIO3ReadMaxAge;
    function AsyncIO3ReadMaxAge(dwCount: DWORD; phServer: POPCHANDLEARRAY; pdwMaxAge: PDWORDARRAY;
                               dwTransactionID: DWORD; out pdwCancelID: DWORD;
                               out ppErrors: PResultList): HResult; stdcall;
    function IOPCAsyncIO3.WriteVQT = AsyncIO3WriteVQT;
    function AsyncIO3WriteVQT(dwCount: DWORD; phServer: POPCHANDLEARRAY; pItemVQT: POPCITEMVQTARRAY;
                             dwTransactionID: DWORD; out pdwCancelID: DWORD;
                             out ppErrors: PResultList): HResult; stdcall;
    function IOPCAsyncIO3.RefreshMaxAge = AsyncIO3RefreshMaxAge;
    function AsyncIO3RefreshMaxAge(dwMaxAge: DWORD; dwTransactionID: DWORD;
                                  out pdwCancelID: DWORD): HResult; stdcall;
    function IOPCAsyncIO3.Read = AsyncIO2Read;
    function IOPCAsyncIO3.Write = AsyncIO2Write;

    { IOPCItemDeadbandMgt }
    function SetItemDeadband(dwCount: DWORD; phServer: POPCHANDLEARRAY; pPercentDeadband: PSingleArray;
                            out ppErrors: PResultList): HResult; stdcall;
    function GetItemDeadband(dwCount: DWORD; phServer: POPCHANDLEARRAY; out ppPercentDeadband: PSingleArray;
                            out ppErrors: PResultList): HResult; stdcall;
    function ClearItemDeadband(dwCount: DWORD; phServer: POPCHANDLEARRAY;
                               out ppErrors: PResultList): HResult; stdcall;

    { IGroupItemList }
    function IGroupItemList.List = GetGroupItemList;
    function GetGroupItemList: TGroupItemList;

    { IConnectionPointContainer }
    function EnumConnectionPoints(
      out enumconn: IEnumConnectionPoints): HResult; stdcall;
    function FindConnectionPoint(const iid: TIID;
      out cp: IConnectionPoint): HResult; stdcall;

    { IDataObject }
    function GetData(const formatetcIn: TFormatEtc; out medium: TStgMedium):
      HResult; stdcall;
    function GetDataHere(const formatetc: TFormatEtc; out medium: TStgMedium):
      HResult; stdcall;
    function QueryGetData(const formatetc: TFormatEtc): HResult;
      stdcall;
    function GetCanonicalFormatEtc(const formatetc: TFormatEtc;
      out formatetcOut: TFormatEtc): HResult; stdcall;
    function SetData(const formatetc: TFormatEtc; var medium: TStgMedium;
      fRelease: BOOL): HResult; stdcall;
    function EnumFormatEtc(dwDirection: Longint; out enumFormatEtc:
      IEnumFormatEtc): HResult; stdcall;
    function DAdvise(const formatetc: TFormatEtc; advf: Longint;
      const advSink: IAdviseSink; out dwConnection: Longint): HResult; stdcall;
    function DUnadvise(dwConnection: Longint): HResult; stdcall;
    function EnumDAdvise(out enumAdvise: IEnumStatData): HResult;
      stdcall;

   { IOPCAsyncIO }
    function IOPCAsyncIO.Read = AsyncIORead;
    function AsyncIORead(dwConnection: DWORD; dwSource: OPCDATASOURCE;
      dwCount: DWORD; phServer: POPCHANDLEARRAY;
      out pTransactionID: DWORD; out ppErrors: PResultList): HResult; stdcall;
    function IOPCAsyncIO.Write = AsyncIOWrite;
    function AsyncIOWrite(dwConnection: DWORD; dwCount: DWORD;
      phServer: POPCHANDLEARRAY; pItemValues: POleVariantArray;
      out pTransactionID: DWORD; out ppErrors: PResultList): HResult; stdcall;
    function IOPCAsyncIO.Refresh = AsyncIORefresh;
    function AsyncIORefresh(dwConnection: DWORD; dwSource: OPCDATASOURCE;
      out pTransactionID: DWORD): HResult; stdcall;
    function IOPCAsyncIO.Cancel = Cancel2;
    procedure ForceDisconnect; {cf 1.15.5.1}
  public
    function ItemCount: Integer; override;
    function Item(i: Integer): TGroupItemInfo; override;
    constructor Create( aServer: TClientInfo {IServerInternal};
                        const aName: String;
                        aActive: Boolean;
                        aUpdateRate: DWORD;
                        ahClientGroup: OPCHANDLE;
                        aTimeBias: Longint;
                        aPercentDeadband: Single;
                        aLCID: DWORD);
    destructor Destroy; override;
  end;

  TGroupKeepAlive = class
  private
    Group : TGroupImpl;
  public
    KeepAlive: DWORD;
    constructor Create(AGroup: TGroupImpl; AKeepAlive : DWORD);
    destructor Destroy; override;
    procedure RestartTimer;
    procedure Tick;
  end;

  TUnkEnumerator = class(TInterfacedObject, IEnumUnknown)
    FIndex: Integer;
    FList: TInterfaceList;
    constructor Create;
    constructor CreateGroupEnumerator( GroupList: TGroupList;
                                       Scope: OPCENUMSCOPE);
    constructor CreateClone(aList: TInterfaceList);
    destructor Destroy; override;
    function Count: Integer;
    function Next(celt: Longint; out elt;
      pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumUnknown): HResult; stdcall;
  end;

  TStringsEnumerator = class(TInterfacedObject, IEnumString)
    FIndex: Integer;
    FStrings: TStrings;
    constructor Create;
    constructor CreateOpcItemsEnumerator(
      ItemServer: TOpcItemServer;
      List: TItemIdList;
      BrowseType: OPCBROWSETYPE;
      const FilterCriterion: String;
      DataTypeFilter: TVarType;
      AccessRightsFilter: DWORD);
    constructor CreateClone(aStrings: TStrings);
    constructor CreateGroupEnumerator(
      GroupList: TGroupList;
      Scope: OPCENUMSCOPE);
    destructor Destroy; override;
    function Count: Integer;
    function Next(celt: Longint; out elt;
      pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumString): HResult; stdcall;
  end;

  {Group reference counting is very complicated.
    See Documentation for IOPCServer.RemoveGroup. I am not
    sure I have implemented this properly &&&}
  TServerImpl = class(TClientInfo, IOPCServer, IOPCCommon,
                     IOPCBrowseServerAddressSpace,
                     IConnectionPointContainer, IOPCItemProperties,
                     {DA3 support} IOPCBrowse, IOPCItemIO)
  private
    FBrowsePos: TItemIdList;
    FGroupList: TGroupList;
    FConnectionPoint: TConnectionPoint;

    {OPC_STATUS_RUNNING
       The server is running normally. This is the usual state for a server
     OPC_STATUS_FAILED
       A vendor specific fatal error has occurred within the server. The server
       is no longer functioning. The recovery procedure from this situation is
       vendor specific. An error code of E_FAIL should generally be returned
       from any other server method.
     OPC_STATUS_NOCONFIG
       The server is running but has no configuration information loaded and
       thus cannot function normally. Note this state implies that the server
       needs configuration information in order to function. Servers which do
       not require configuration information should not return this state.
    OPC_STATUS_SUSPENDED
       The server has been temporarily suspended via some vendor specific method
       and is not getting or sending data. Note that Quality will be returned as
       OPC_QUALITY_OUT_OF_SERVICE.
    OPC_STATUS_TEST
       The server is in Test Mode. The outputs are disconnected from the real
       hardware but the server will otherwise behave normally. Inputs may be real
       or may be simulated depending on the vendor implementation. Quality will
       generally be returned normally.}

    procedure SinkConnect(const Sink: IUnknown; Connecting: Boolean);
    function SetLastUpdate: TFileTime;

    { IOPCServer }
    function AddGroup(szName: POleStr; bActive: BOOL; dwRequestedUpdateRate: DWORD;
                      hClientGroup: OPCHANDLE; pTimeBias: PLongint; pPercentDeadband: PSingle;
                      dwLCID: DWORD;
                      out phServerGroup: OPCHANDLE;
                      out pRevisedUpdateRate: DWORD;
                      const riid: TIID;
                      out ppUnk: IUnknown): HResult; stdcall;
    function IOPCServer.GetErrorString = OpcServerGetErrorString;
    function OpcServerGetErrorString(dwError: HResult; dwLocale: TLCID; out ppString: POleStr): HResult; stdcall;
    function GetGroupByName(szName: POleStr; const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
    function GetStatus(out ppServerStatus: POPCSERVERSTATUS): HResult; stdcall;
    function RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
    function CreateGroupEnumerator(dwScope: OPCENUMSCOPE; const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;

    { IOPCCommon }
    function SetLocaleID(dwLcid: TLCID): HResult; stdcall;
    function GetLocaleID(out pdwLcid: TLCID): HResult; stdcall;
    function QueryAvailableLocaleIDs(out pdwCount: UINT; out pdwLcid: PLCIDARRAY): HResult; stdcall;

    function IOPCCommon.GetErrorString = OPCCommonGetErrorString;
    function OPCCommonGetErrorString(dwError: HResult; out ppString: POleStr): HResult; stdcall;
    function SetClientName(szName: POleStr): HResult; stdcall;

    { IOPCBrowseServerAddressSpace }
    function QueryOrganization(out pNameSpaceType: OPCNAMESPACETYPE): HResult; stdcall;
    function ChangeBrowsePosition(dwBrowseDirection: OPCBROWSEDIRECTION;
                                 szString: POleStr): HResult; stdcall;
    function BrowseOPCItemIDs(dwBrowseFilterType: OPCBROWSETYPE; szFilterCriteria: POleStr;
                             vtDataTypeFilter: TVarType; dwAccessRightsFilter: DWORD;
                             out ppIEnumString: IEnumString): HResult; stdcall;
    function GetItemID(szItemDataID: POleStr; out szItemID: POleStr): HResult; stdcall;
    function BrowseAccessPaths(szItemID: POleStr; out ppIEnumString: IEnumString): HResult; stdcall;


    {these functions are used to save/restore browsing context before and after
    invalidating the namespace. Note that a dynamic namespace is not discussed
    in the spec but it is provided for good practical reasons. One of the
    inevitable consequences is that browsing context could be unexpectedly lost}
    function ClearBrowsingContext: String;
    procedure RestoreBrowsingContext(const S: String);

    { IOPCItemProperties }
    function QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                          out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                          out ppvtDataTypes:PVarTypeList):HResult; stdcall;
    function GetItemProperties(szItemID:POleStr; dwCount:DWORD;
           pdwPropertyIDs:PDWORDARRAY; out ppvData:POleVariantArray; out ppErrors:PResultList):HResult; stdcall;
    function LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
           out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;

    { IConnectionPointContainer }
    function EnumConnectionPoints(
      out enumconn: IEnumConnectionPoints): HResult; stdcall;
    function FindConnectionPoint(const iid: TIID;
      out cp: IConnectionPoint): HResult; stdcall;

    { IOPCBrowse }
    function FillProperties(szItemID: POleStr; bReturnPropertyValues: BOOL; dwPropertyCount: DWORD;
                           pdwPropertyIDs: PDWORDARRAY; var ItemProperties: OPCITEMPROPERTIES): HResult;
    function GetProperties(dwItemCount: DWORD; pszItemIDs: POleStrList; bReturnPropertyValues: BOOL;
                          dwPropertyCount: DWORD; pdwPropertyIDs: PDWORDARRAY;
                          out ppItemProperties: POPCITEMPROPERTIESARRAY): HResult; stdcall;
    function Browse(szItemID: POleStr; var pszContinuationPoint: POleStr; dwMaxElementsReturned: DWORD;
                   dwBrowseFilter: OPCBROWSEFILTER; szElementNameFilter: POleStr; szVendorFilter: POleStr;
                   bReturnAllProperties: BOOL; bReturnPropertyValues: BOOL; dwPropertyCount: DWORD;
                   pdwPropertyIDs: PDWORDARRAY; out pbMoreElements: BOOL; out pdwCount: DWORD;
                   out ppBrowseElements: POPCBROWSEELEMENTARRAY): HResult; stdcall;

    { IOPCItemIO }
    function Read(dwCount: DWORD; pszItemIDs: POleStrList; pdwMaxAge: PDWORDARRAY;
                 out ppvValues: POleVariantArray; out ppwQualities: PWordArray;
                 out ppftTimeStamps: PFileTimeArray; out ppErrors: PResultList): HResult; stdcall;
    function WriteVQT(dwCount: DWORD; pszItemIDs: POleStrList; pItemVQT: POPCITEMVQTARRAY;
                     out ppErrors: PResultList): HResult; stdcall;

  public
    procedure ForceDisconnect; override;
    function ObjQueryInterface(const IID: TGUID; out Obj): HResult; override;
    procedure Initialize; override;
    destructor Destroy; override;
    function GroupCount: Integer; override;
    function Group(i: Integer): TGroupInfo; override;
  end;

  TEnumerateCP = class(TInterfacedObject, IEnumConnectionPoints) {cf 1.01.5}
    CP: IConnectionPoint;
    FIndex: Integer;
    function Next(celt: Longint; out elt;
      pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out Enum: IEnumConnectionPoints): HResult;
      stdcall;
    constructor Create(aCP: IConnectionPoint; aIndex: Integer);
  end;

const
  NDa1Advise: array[TDa1Format] of String = (
    'OPCSTMFORMATDATA',
    'OPCSTMFORMATDATATIME',
    'OPCSTMFORMATWRITECOMPLETE');

  StdItemPropCount = 6;
  StdItemPropID: array[0..StdItemPropCount-1] of DWORD = (
    OPC_PROPERTY_DATATYPE,
    OPC_PROPERTY_VALUE,
    OPC_PROPERTY_QUALITY,
    OPC_PROPERTY_TIMESTAMP,
    OPC_PROPERTY_ACCESS_RIGHTS,
    OPC_PROPERTY_SCAN_RATE);
  StdItemPropDesc: array[0..StdItemPropCount-1] of String = (
    OPC_PROPERTY_DESC_DATATYPE,
    OPC_PROPERTY_DESC_VALUE,
    OPC_PROPERTY_DESC_QUALITY,
    OPC_PROPERTY_DESC_TIMESTAMP,
    OPC_PROPERTY_DESC_ACCESS_RIGHTS,
    OPC_PROPERTY_DESC_SCAN_RATE);
  StdItemPropType: array[0..StdItemPropCount-1] of TVarType = (
    VT_I2,
    VT_VARIANT,  {dependent on item id}
    VT_I2,
    VT_DATE,
    VT_I4,
    VT_R4);

  MAX_UPDATE_RATE = 20;  {milliseconds}


{$IFDEF Evaluation}
  MaxEvalItems = 40;
  MaxEvalClients = 3;
{$ENDIF}

{$IFDEF GLD}
procedure SendDebug(const S: String);
begin
  OutputDebugString(PChar(S))
end;
{$ENDIF}


resourcestring
  SNoItemServer = 'OPCSERV: No item server';
  SItemServerOnlyOne = 'OPCSERV: Only one instance of namespace allowed';
  SCouldNotAllocateOPCWindow = 'OPCSERV: Could not allocate Opc Window';
  SUnknownError = 'Unknown error code %.8x';

  SCannotCallNewBranchOnFlatSpace = 'Cannot call NewBranch in flat address space';
  SInvalidBranchName = '%s is an invalid branch name';
  SDuplicateItemId = 'Duplicate Item Id %s';

var
  GOpcItemServer: TOpcItemServer;
  GDa1Format: array[TDa1Format] of TClipFormat = (0, 0, 0);

function OpcAccessRights(AccessRights: TAccessRights): DWORD;
{OPC_READABLE = 1 and OPC_WRITABLE = 2 so
  AccessRights are binary compatible}
begin
  Result:= 0;
  Move(AccessRights, Result, 1)
end;

function NativeAccessRights(OpcRights: DWORD): TAccessRights;
begin
  Result:= [];
  Move(OpcRights, Result, 1)
end;

procedure CheckRequestedVarType(VarType: Integer);
begin
  if not IsSupportedVarType(VarType) then
    raise EOpcError.Create(OPC_E_BADTYPE)
end;

procedure TestParamRange(Param: Integer; Low, High: Integer);
begin
  if (Param < Low) or (Param > High) then
    raise EOpcError.Create(E_INVALIDARG)
end;

function OpcItemServer: TOpcItemServer;
begin
  if not Assigned(GOpcItemServer) then
    raise EOpcServer.CreateRes(@SNoItemServer);
  Result:= GOpcItemServer
end;

function LockServerList: TList;
begin
  Result:= OpcItemServer.FServerList.LockList
end;

procedure UnlockServerList;
begin
  OpcItemServer.FServerList.UnlockList
end;


function ServerItemList: TServerItemList;
begin
  Result:= TServerItemList(OpcItemServer.FItemList)
end;

type
  TOpcServerFactory = class(TComObjectFactory)
  private
    FVendorName: String;
    FVersionIndependentProgID : String;
  public
    procedure UpdateRegistry(Register: Boolean); override;
    constructor Create(ComServer: TComServerObject; ComClass: TComClass;
      const ClassID: TGUID; const ClassName, VersionIndependentProgID, Description, VendorName: string);
  end;

procedure TOpcServerFactory.UpdateRegistry(Register: Boolean);

  procedure RemoveBrowserKey;
  var
    Cr: ICatRegister;
{   Info: array[0..1] of TCategoryInfo; }
    Catid: array[0..2] of TGUID;
  begin
    CoInitialize(nil);
    OleCheck(CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
        CLSCTX_INPROC_SERVER, ICatRegister, Cr));
    Catid[0]:= CATID_OPCDAServer30;
    Catid[1]:= CATID_OPCDAServer20;
    Catid[2]:= CATID_OPCDAServer10;
    Cr.UnRegisterClassImplCategories(ClassID, 3, @Catid);
    {ignore the error if this fails. It probably just means that categories
    are already unregistered}

    {I have observed that unregistering categories removes them,
     even if there are other classes on the system that are still referencing them
     I don't know the correct way to avoid this problem. For now, I just leave the
     categories registered on the system after deinstallation of the server. This
     seems better than possibly breaking somebody else's server

    with Info[0] do
    begin
      catid:= CATID_OPCDAServer20;
      lcid:= GetUserDefaultLCID;
      StringToWideChar(CATID_OPCDAServer20Desc, szDescription, Length(szDescription)-1)
    end;
    with Info[1] do
    begin
      catid:= CATID_OPCDAServer10;
      lcid:= GetUserDefaultLCID;
      StringToWideChar(CATID_OPCDAServer10Desc, szDescription, Length(szDescription)-1)
    end;
    OleCheck(Cr.UnRegisterCategories(2, @Info));
}
  end;

  procedure RemoveDA1BrowserKey;
  var
    OpcKey: String;
  begin
    {blit old registration key if possible}
    OpcKey:= ProgID + '\Opc';
    ComObj.DeleteRegKey(OpcKey + '\Vendor');
    ComObj.DeleteRegKey(OpcKey)
  end;

  procedure InstallBrowserKey;
  var
    Cr: ICatRegister;
    Info: array[0..2] of TCategoryInfo;
    Catid: array[0..2] of TGUID;
  begin
    CoInitialize(nil);
    OleCheck(CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
        CLSCTX_INPROC_SERVER, ICatRegister, Cr));
    with Info[0] do
    begin
      catid:= CATID_OPCDAServer30;
      lcid:= GetUserDefaultLCID;
      StringToWideChar(CATID_OPCDAServer30Desc, szDescription, Length(szDescription)-1)
    end;
    with Info[1] do
    begin
      catid:= CATID_OPCDAServer20;
      lcid:= GetUserDefaultLCID;
      StringToWideChar(CATID_OPCDAServer20Desc, szDescription, Length(szDescription)-1)
    end;
    with Info[2] do
    begin
      catid:= CATID_OPCDAServer10;
      lcid:= GetUserDefaultLCID;
      StringToWideChar(CATID_OPCDAServer10Desc, szDescription, Length(szDescription)-1)
    end;
    OleCheck(Cr.RegisterCategories(3, @Info));
    Catid[0]:= CATID_OPCDAServer30;
    Catid[1]:= CATID_OPCDAServer20;
    Catid[2]:= CATID_OPCDAServer10;
    OleCheck(Cr.RegisterClassImplCategories(ClassID, 3, @Catid))
  end;

  procedure InstallDA1BrowserKey;
  var
    OpcKey: String;
    Key: HKey;
  begin
    try
      OpcKey:= ProgID + '\Opc';
      if RegOpenKey(HKEY_CLASSES_ROOT, PChar(OpcKey), Key) = ERROR_SUCCESS then
      begin
        RegCloseKey(Key)
      end else
      begin
        ComObj.CreateRegKey(OpcKey, '', '');
        OpcKey:= OpcKey + '\Vendor';
        ComObj.CreateRegKey(OpcKey, '', FVendorName)
      end
    except
      on EOleRegistrationError do ; {why &&&}
    end
  end;

begin
  inherited UpdateRegistry(Register);
  ComObj.CreateRegKey('AppID\'+GUIDToString(ClassID), '', Description);
  ComObj.CreateRegKey('AppID\'+ExtractFileName(ParamStr(0)),'AppId', GUIDToString(ClassID));
  ComObj.CreateRegKey('CLSID\' + GUIDToString(ClassID) + '\VersionIndependentProgID', '', ComServer.ServerName + '.' + FVersionIndependentProgID);
  if Register then
  begin
    InstallBrowserKey;
    InstallDA1BrowserKey
  end else
  begin
    RemoveBrowserKey;
    RemoveDA1BrowserKey
  end
end;

constructor TOpcServerFactory.Create(ComServer: TComServerObject; ComClass: TComClass;
  const ClassID: TGUID; const ClassName, VersionIndependentProgID, Description, VendorName: string);
begin
  inherited Create(ComServer, ComClass, ClassID, ClassName, Description,
    ciMultiInstance, tmApartment);
  FVendorName:= VendorName;
  FVersionIndependentProgID := VersionIndependentProgID;
end;


procedure RegisterOPCServer(const ServerGUID: TGUID;
                               ServerVersion: Integer;
                            const ServerDesc: String;
                          const ServerVendor: String;
                              aOpcItemServer: TOpcItemServer);
var
  i: TDa1Format;
  aServerDesc: string;
begin
  for i:= Low(TDa1Format) to High(TDa1Format) do
  begin
    GDa1Format[i]:= RegisterClipboardFormat(PChar(NDa1Advise[i]));
    if GDa1Format[i] = 0 then
      raise EOpcServer.CreateFmt('Could not register %s clip format',
        [NDa1Advise[i]])
  end;
//  GOpcItemServer:= aOpcItemServer;
(*
  GAlwaysAllocPpErrors:= soAlwaysAllocateErrorArrays in aOpcItemServer.Options; {cf. 1.12.1}
  GNonCompliantDA1Stream:= soIncludeNullInDA1ByteCount in aOpcItemServer.Options;
  GStrictHandles:= soStrictHandleValidation in aOpcItemServer.Options; {cf 1.14.10}
  GHierarchical:= soHierarchicalBrowsing in aOpcItemServer.Options; {cf 1.14.18}
*)
  if ServerDesc <> '' then
    aServerDesc:= ServerDesc
  else
    aServerDesc:= ComServer.ServerName;
  aOpcItemServer.FVendorInfo:= ServerVendor + ' ' + aServerDesc;
  TOpcServerFactory.Create(ComServer,
    TServerImpl, ServerGUID,
    aOpcItemServer.GetServerID(ServerVersion),
    aOpcItemServer.GetServerVersionIndependentID,
    aServerDesc, ServerVendor)
end;

procedure CheckCount(dwCount: DWORD; InputArray: Pointer);
begin
  if (dwCount = 0) or (InputArray = nil) then
    raise EOpcError.Create(E_INVALIDARG);
end;

function CheckAllocation(Size: DWORD): Pointer;
begin
  Result:= CoTaskMemAlloc(Size);
  if Result = nil then
    raise EOpcError.Create(E_OUTOFMEMORY)
end;

function ZeroAllocation(Size: DWORD): Pointer;
begin
  Result:= CheckAllocation(Size);
  FillChar(Result^, Size, 0)
end;

procedure FreeAndNull(var P);
begin
  if Assigned(Pointer(P)) then
  begin
    CoTaskMemFree(Pointer(P));
    Pointer(P):= nil
  end
end;

procedure InitItemID(const ItemDef: OPCITEMDEF; var ItemID: String);
{var
  AccessPath: WideString; }
begin
{  AccessPath:= ItemDef.szAccessPath;
  if AccessPath <> '' then
    raise EOpcError.Create(OPC_E_UNKNOWNPATH); Access paths not supported
  section 6.2 of the spec states:
    Servers which do not support access paths will completely ignore
    any passed access path (and will not treat this as an error by the client)}
  ItemID:= ItemDef.szItemID;
  if ItemID = '' then
    raise EOpcError.Create(OPC_E_INVALIDITEMID);
  if not IsSupportedVarType(ItemDef.vtRequestedDataType) then
    raise EOpcError.Create(OPC_E_BADTYPE)
end;

procedure ReadItemResult( ItemResult: TServerItemRef; var Res: OPCITEMRESULT);
begin
  FillChar(Res, SizeOf(Res), 0);
  Res.vtCanonicalDataType:= ItemResult.CanonicalDataType;
  Res.dwAccessRights:= ItemResult.AccessRights
end;

const
  CM_ASYNCTASK = $8FFF;     {wParam: TAsyncTask}
  CM_NOTIFICATION = $8FFE;
    nfClientConnect = 1;
    nfClientDisconnect = 2;
    nfClientSetName = 3;
    nfAddGroup = 4;
    nfRemoveGroup = 5;
    nfAddItem = 6;
    nfRemoveItem = 7;

function OpcWndProc(Window: HWND; Message, wParam, lParam: Longint): Longint; stdcall;
var
  Obj: TObject absolute wParam;
  Task: TAsyncTask absolute wParam;
  Group: TGroupImpl absolute wParam;
  KeepAlive: TGroupKeepAlive absolute wParam;
begin
  try
    case Message of
      CM_ASYNCTASK:
      begin
        if Assigned(Task) then
        begin
          Task.DoProcess; {cf 1.15.5.2}
          Task.Destroy
        end;
        Result:= 0
      end;
      CM_NOTIFICATION:
      begin
        if Assigned(GOpcItemServer) then
        with GOpcItemServer do
        case wParam of
          nfClientConnect: OnClientConnect(TClientInfo(lParam));
          nfClientDisconnect: OnClientDisconnect(TClientInfo(lParam));
          nfClientSetName: OnClientSetName(TClientInfo(lParam));
          nfAddGroup: OnAddGroup(TGroupInfo(lParam));
          nfRemoveGroup: OnRemoveGroup(TGroupInfo(lParam));
          nfAddItem: OnAddItem(TGroupItemInfo(lParam));
          nfRemoveItem: OnRemoveItem(TGroupItemInfo(lParam));
        end;
        Result:= 0
      end;
      WM_TIMER:
      begin
        if Obj is TGroupImpl then
          Group.Tick;
        if Obj is TGroupKeepAlive then
          KeepAlive.Tick;
        Result:= 0
      end
    else
      Result:= DefWindowProc(Window, Message, wParam, lParam);
    end;
  except
    GOpcItemServer.UnhandledException(ExceptObject); {cf 1.14.22}
    Result:= 1
  end
end;

{$IFDEF SLOWCALLBACKS}
const
  TestAsyncWait = 100; {ms}

procedure TimerProc(hw: HWND; uMsg, idEvent, dwTime: Cardinal); stdcall;
var
  Task: TAsyncTask absolute idEvent;
begin
  KillTimer(hw, idEvent);
  Task.Free
end;

procedure TestPostMessage(hw: HWND; Msg: Cardinal; wParam, lParam: Integer);
{create a timer}
begin
  SetTimer(hw, wParam, TestAsyncWait, @TimerProc)
end;
{$ENDIF}

var
  OpcWindow: HWND = 0;

procedure FreeOpcWindow;
begin
  DestroyWindow(OpcWindow)
end;

procedure AllocateWindow;
var
  OpcWndClass: TWndClass;
begin
  FillChar(OpcWndClass, SizeOf(OpcWndClass), 0);
  OpcWndClass.hInstance:= HInstance;
  OpcWndClass.lpfnWndProc:= @OpcWndProc;
  OpcWndClass.lpszClassName:= 'PrelOpcServerWindow';
  Windows.RegisterClass(OpcWndClass);
  OpcWindow:= CreateWindow(OpcWndClass.lpszClassName, '', 0,
    0, 0, 0, 0, 0, 0, HInstance, nil)
end;

{ TConnectionPoint }

constructor TConnectionPoint.Create(const Container: IUnknown;
  const IID: TGUID;
  OnConnect: TConnectEvent);
begin
  inherited Create;
  FContainer:= Pointer(Container);
  FIID:= IID;
  FOnConnect:= OnConnect
end;

destructor TConnectionPoint.Destroy;
begin
  FSink:= nil;
  inherited Destroy
end;

{ TConnectionPoint.IConnectionPoint }

function TConnectionPoint.GetConnectionInterface(out iid: TIID): HResult;
begin
  iid:= FIID;
  Result:= S_OK;
end;

function TConnectionPoint.GetConnectionPointContainer(
  out cpc: IConnectionPointContainer): HResult;
begin
  cpc:= IUnknown(FContainer) as IConnectionPointContainer;
  Result:= S_OK
end;

function TConnectionPoint.Advise(const unkSink: IUnknown;
  out dwCookie: Longint): HResult;
begin
  if Assigned(FSink) then
  begin
    Result:= CONNECT_E_CANNOTCONNECT
  end else
  begin
    if Assigned(FOnConnect) then
      FOnConnect(unkSink, True);
    dwCookie:= 1;
    FSink:= unkSink;
    Result:= S_OK
  end
end;

function TConnectionPoint.Unadvise(dwCookie: Longint): HResult;
begin
  if not Assigned(FSink) then
  begin
    Result:= CONNECT_E_NOCONNECTION
  end else
  begin
    if Assigned(FOnConnect) then
      FOnConnect(FSink, False);
    FSink:= nil;
    Result:= S_OK
  end
end;

function TConnectionPoint.EnumConnections(out enumconn: IEnumConnections): HResult;
begin
  Result:= E_NOTIMPL
end;

function TConnectionPoint._AddRef: Integer;
begin
  Result:= IUnknown(FContainer)._AddRef
end;

function TConnectionPoint._Release: Integer;
begin
  Result:= IUnknown(FContainer)._Release
end;

function TConnectionPoint.QueryInterface(const IID: TGUID;
  out Obj): HResult;
begin
  if GetInterface(IID, Obj) then
    Result:= S_OK
  else
    Result:= E_NOINTERFACE
end;

{ TOpcItemServer }

constructor TOpcItemServer.Create;
begin
  if Assigned(GOpcItemServer) then
    raise EOpcServer.CreateRes(@SItemServerOnlyOne);
  inherited Create;
  GOpcItemServer:= Self;
  AllocateWindow;
  if OpcWindow = 0 then
    raise EOpcServer.CreateRes(@SCouldNotAllocateOPCWindow);
  FItemList:= TServerItemList.Create;
  FServerList:= TThreadList.Create;
  GetTimestamp(FStartupTime) {cf 1.14.28}
end;

destructor TOpcItemServer.Destroy;
begin
  FreeOpcWindow;
  FItemList.Free;
  FServerList.Free;
  FRootNode.Free;
  GOpcItemServer:= nil;
  inherited Destroy
end;

function TOpcItemServer.MaxUpdateRate: Cardinal;
begin
  Result:= MAX_UPDATE_RATE
end;

function TOpcItemServer.SubscribeToItem(ItemHandle: TItemHandle;
  UpdateEvent: TSubscriptionEvent): Boolean;
begin
  Result:= false
end;

procedure TOpcItemServer.UnsubscribeToItem(
  ItemHandle: TItemHandle);
begin
end;

{$IFDEF NoMasks}
function MatchesMask(const ItemId, Mask: string): Boolean;
begin
  Result:= true
end;
{$ENDIF}

function TOpcItemServer.FilterFunction(const FilterMask: String;
                            FilterAccessRights: TAccessRights;
                            FilterVarType: TVarType;
                            const ItemID: String;
                            ItemAccessRights: TAccessRights;
                            ItemVarType: TVarType): Boolean;
begin
  try
    Result:= ((FilterMask = '') or MatchesMask(ItemID, FilterMask)) and
       ((FilterAccessRights = []) or ((FilterAccessRights * ItemAccessRights) <> [])) and
       ((FilterVarType = VT_EMPTY) or (FilterVarType = ItemVarType));
  except
    on Exception do
      raise EOpcError.Create(OPC_E_INVALIDFILTER)
  end
end;

procedure TOpcItemServer.ReleaseHandle(ItemHandle: TItemHandle);
begin
end;

function TOpcItemServer.ClientCount: Integer;
begin
  with FServerList.LockList do
  try
    Result:= Count
  finally
    FServerList.UnlockList
  end
end;

function TOpcItemServer.ClientInfo(i: Integer): TClientInfo;
begin
  with FServerList.LockList do
  try
    Result:= TClientInfo(Items[i])
  finally
    FServerList.UnlockList
  end
end;

function TOpcItemServer.ShutdownRequest(const Reason: String;
  ClientNo: Integer): Integer;
var
  i: Integer;
  CC: Integer;
  CopyList: array of TServerImpl;
begin
  {this is a bit tricky. Clients may shutdown while this
   is executing which will cause the list to shrink. We must
   therefore make a copy of the list then interate over that}
  Result:= 0;
  if ClientNo = -1 then
  begin
    with FServerList.LockList do
    try
      CC:= Count;
      SetLength(CopyList, CC);
      for i:= 0 to CC - 1 do
        CopyList[i]:= TServerImpl(Items[i])
    finally
      FServerList.UnlockList
    end;
    for i:= 0 to CC - 1 do
    if CopyList[i].ShutdownRequest(Reason) then
      Inc(Result)
  end else
  begin
    if ClientInfo(ClientNo).ShutdownRequest(Reason) then
      Inc(Result)
  end
end;

procedure TOpcItemServer.OnAddGroup(Group: TGroupInfo);
begin
end;

procedure TOpcItemServer.OnAddItem(Item: TGroupItemInfo);
begin
end;

procedure TOpcItemServer.OnItemValueChange(Item: TGroupItemInfo);
begin
end;

procedure TOpcItemServer.OnRemoveItem(Item: TGroupItemInfo);
begin
end;

procedure TOpcItemServer.OnClientConnect(aServer: TClientInfo);
begin
end;

procedure TOpcItemServer.OnClientDisconnect(aServer: TClientInfo);
begin
end;

procedure TOpcItemServer.OnRemoveGroup(Group: TGroupInfo);
begin
end;

procedure TOpcItemServer.OnClientSetName(aServer: TClientInfo);
begin
end;


function TOpcItemServer.GetItemInfo(const ItemID: String; var AccessPath: String;
  var AccessRights: TAccessRights): Integer;
begin
  raise EOpcError.Create(OPC_E_INVALIDITEMID)
end;

function TOpcItemServer.GetExtendedItemInfo(const ItemID: String;
                        var AccessPath: String;
                        var AccessRights: TAccessRights;
                        var EUInfo: IEUInfo;
                        var ItemProperties: IItemProperties): Integer;
begin
  Result:= GetItemInfo(ItemID, AccessPath, AccessRights)
end;

procedure TOpcItemServer.ListItemIDs(List: TItemIDList);
begin
end;

function TOpcItemServer.GetItemVQT(ItemHandle: TItemHandle;
  var Quality: Word; var Timestamp: TFileTime): OleVariant;
begin
  GOpcItemServer.GetTimestamp(Timestamp);
  Result := GetItemValue(ItemHandle, Quality);
end;

function TOpcItemServer.GetItemValue(ItemHandle: TItemHandle;
  var Quality: Word): OleVariant;
begin
  raise Exception.Create('Called abstract method TOpcItemServer.GetItemValue');
end;

procedure TOpcItemServer.SetItemVQT(ItemHandle: TItemHandle; const ValueVQT: OPCITEMVQT);
begin
  SetItemValue(ItemHandle, ValueVQT.vDataValue);

  if ValueVQT.bQualitySpecified then
   SetItemQuality(ItemHandle, ValueVQT.wQuality);

  if ValueVQT.bTimeStampSpecified then
   SetItemTimestamp(ItemHandle, ValueVQT.ftTimeStamp);
end;

procedure TOpcItemServer.SetItemQuality(ItemHandle: TItemHandle; const Quality: Word);
begin
end;

procedure TOpcItemServer.SetItemTimestamp(ItemHandle: TItemHandle; const Timestamp: TFileTime);
begin
end;

function TOpcItemServer.Options: TServerOptions;
begin
  Result:= [soAlwaysAllocateErrorArrays]
end;

function TOpcItemServer.GetServerID(Version: Integer): string;
begin
  FmtStr(Result, '%s.%d', [GetServerVersionIndependentID, Version])
end;

function TOpcItemServer.GetServerVersionIndependentID: string;
begin
  Result := ClassName;
end;


function TOpcItemServer.GetErrorString(Code: HResult; LCID: DWORD): String;
var
  Buf: array[0..255] of Char;
begin
  Result:= '';
  if not StdOpcErrorToStr(Code, Result) then
  begin
    if  FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_ARGUMENT_ARRAY,
          nil, DWORD(Code), LCID and $FFFF, Buf, SizeOf(Buf), nil) > 0 then
      Result:= Buf
    else
      raise EOpcError.Create(E_INVALIDARG)
  end
end;

procedure TOpcItemServer.GroupCallbackError(Exception: EOleSysError;
  Group: TGroupInfo; Call: TGroupCallback);
begin
{$IFDEF GLD}
  SendDebug(Format('Group callback error %s',
    [GetEnumName(TypeInfo(TGroupCallback), Integer(Call))]))
{$ENDIF}
end;

procedure TOpcItemServer.ClientCallbackError(Exception: EOleSysError;
  Server: TClientInfo; Call: TClientCallback);
begin
{$IFDEF GLD}
  SendDebug(Format('Client callback error %s',
    [GetEnumName(TypeInfo(TClientCallback), Integer(Call))]))
{$ENDIF}
end;


{  Property support
function TOpcItemServer.GetItemIdFromPropertyInfo(const ParentId: string;
  Pid: Integer): string;
begin
  Result:= ''
end;

function TOpcItemServer.GetPropertyInfo(
  const ItemId: string): TItemProperties;
begin
  Result:= nil
end;

procedure TOpcItemServer.GetPropertyInfoFromItemId(const ItemId: string;
  var ParentId: string; var Pid: Integer);
begin
  ParentId:= ItemId;
  Pid:= 0
end;
}

function TOpcItemServer.PathDelimiter: Char;
begin
  Result:= '.'
end;

function TOpcItemServer.PropertyDelimiter: Char;
begin
  Result:= '.'
end;

procedure TOpcItemServer.InvalidateNamespace(KeepExisting: Boolean);
var
  sl: TList;
  bc: TStrings;
  i: Integer;
  si: TServerImpl; {cf X}
begin
  bc:= TStringList.Create;
  sl:= LockServerList;
  try
    {save browsing context}
    for i:= 0 to sl.Count - 1 do
    begin
      si:= TServerImpl(sl[i]);  {cf 1.15.6.2}
      bc.AddObject(si.ClearBrowsingContext, si)
    end;

    if Assigned(FRootNode) then
    begin
      if not KeepExisting then
        FRootNode.Clear
    end else
    begin
      FRootNode:= TItemIdList.Create(nil, '')
    end;
    {this is like a parameter to ItemIdList.AddItemId to ensure that if old
    ids are kept then adding them again does not cause an error}
    FIgnoreDuplicatesInListItemIds:= KeepExisting;
    ListItemIDs(FRootNode);

    {restore browsing context}
    for i:= 0 to bc.Count - 1 do
    begin
      si:= TServerImpl(bc.Objects[i]);  {cf 1.15.6.2}
      si.RestoreBrowsingContext(bc[i])
    end
  finally
    UnlockServerList;
    bc.Free
  end
end;

function TOpcItemServer.RootNode: TItemIdList;
{Note: it is not necessary to lock the namespace here as we may assume that
this function is only called by the main thread. The namespace is only ever
modified in the main thread. Client threads (Com objects) must lock the
namespace to protect against reading a corrupt namespace}
begin
  if not Assigned(FRootNode) then
    InvalidateNamespace(False);
  Result:= FRootNode
end;

procedure TOpcItemServer.InitBrowsing(var CurrentNode: TItemIdList);
begin
  if not Assigned(FRootNode) then
  begin
    InvalidateNamespace(False);
    CurrentNode:= FRootNode
  end else
  if not Assigned(CurrentNode) then
  begin
    CurrentNode:= FRootNode
  end
end;

function TOpcItemServer.GetServerOption(Option: TServerOption): Boolean;
begin
  Result:= Option in Options
end;

procedure TOpcItemServer.ParseItemId(const ItemId: string; Path: TStrings);
var
  Delimiter: Char;
  P, P1: PChar;
  S: string;
begin
  Path.BeginUpdate;
  try
    Path.Clear;
    Delimiter:= GOpcItemServer.PathDelimiter;
    P:= PChar(ItemId);
    P1:= P;
    while P^ <> #0 do
    begin
      if P^ = Delimiter then
      begin
        SetString(S, P1, P - P1);
        Path.Add(S);
        P1:= P + 1
      end;
      Inc(P)
    end;
    SetString(S, P1, P - P1);
    Path.Add(S)
  finally
    Path.EndUpdate
  end
end;

procedure TOpcItemServer.ParseItemId(const ItemId: string; Path: TStrings;
  var PropName: string);
begin
  ParseItemId(ItemId, Path)
end;


procedure TOpcItemServer.UnhandledException(E: TObject);
begin
end;

procedure TOpcItemServer.GetTimestamp(var Timestamp: TFileTime);
begin
  GetSystemTimeAsFileTime(Timestamp) {cf 1.14.28}
end;

{ TStringsEnumerator }

constructor TStringsEnumerator.Create;
begin
  inherited Create;
  FStrings:= TStringList.Create
end;

constructor TStringsEnumerator.CreateOpcItemsEnumerator(
  ItemServer: TOpcItemServer;
  List: TItemIdList;
  BrowseType: OPCBROWSETYPE;
  const FilterCriterion: String;
  DataTypeFilter: TVarType;
  AccessRightsFilter: DWORD);

var
  Ar: TAccessRights;

  procedure AddLeaf(Leaf: TNamespaceItem);
  var
    iid: String;
  begin
    with Leaf do
    begin
      if BrowseType = OPC_FLAT then
        iid:= Path
      else
        iid:= Name;
      if ItemServer.FilterFunction(
        FilterCriterion, Ar, DataTypeFilter, iid, AccessRights, VarType) then
        FStrings.Add(iid)
    end
  end;

  procedure AddBranch(Branch: TItemIdList);
  var
    i: Integer;
    c: TNamespaceNode;
    l: TNamespaceItem absolute c;
    b: TItemIdList absolute c;
  begin
    for i:= 0 to Branch.ChildCount - 1 do
    begin
      c:= Branch.Child(i);
      case BrowseType of
        OPC_BRANCH:
        if (c is TItemIdList) and
          ItemServer.FilterFunction(FilterCriterion, Ar, 0, c.Name, [], 0) then
          FStrings.Add(c.Name);
        OPC_LEAF:
        if c is TNamespaceItem then
          AddLeaf(l);
        OPC_FLAT:
        if c is TNamespaceItem then
          AddLeaf(l)
        else
        if c is TItemIdList then
          AddBranch(b);
      end
    end
  end;

begin
  Create;
  Ar:= NativeAccessRights(AccessRightsFilter);
  AddBranch(List)
end;

constructor TStringsEnumerator.CreateGroupEnumerator(GroupList: TGroupList;
  Scope: OPCENUMSCOPE);
var
  i: Integer;
begin
  Create;
  for i:= 0 to GroupList.Count - 1 do
  with GroupList.Group(i) do
  if InScope(Scope) then
    FStrings.Add(FName)
end;

destructor TStringsEnumerator.Destroy;
begin
  FStrings.Free;
  inherited Destroy
end;

function TStringsEnumerator.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
  i: Integer;
  ItemCount: Integer;
begin
  ItemCount:= FStrings.Count - FIndex;
  if celt > ItemCount then
  begin
    Result:= S_FALSE
  end else
  begin
    ItemCount:= celt;
    Result:= S_OK
  end;
  if Assigned(pceltFetched) then
    pceltFetched^:= ItemCount;
  if ItemCount > 0 then
  begin
    i:= 0;
    while i < ItemCount do
    begin
      TPointerList(elt)[i]:= StringToLPOLESTR(FStrings[FIndex]);
      Inc(FIndex);
      Inc(i)
    end
  end
end;

function TStringsEnumerator.Skip(celt: Longint): HResult;
begin
  if (FIndex + celt) <= FStrings.Count then
  begin
    Inc(FIndex, celt);
    Result:= S_OK
  end else
  begin
    FIndex:= FStrings.Count;
    Result:= S_FALSE
  end
end;

function TStringsEnumerator.Reset: HResult;
begin
  FIndex:= 0;
  Result:= S_OK;
end;

function TStringsEnumerator.Clone(out enm: IEnumString): HResult;
begin
  try
    enm:= TStringsEnumerator.CreateClone(FStrings);
    Result:= S_OK
  except
    Result:= E_UNEXPECTED
  end;
end;

constructor TStringsEnumerator.CreateClone(aStrings: TStrings);
begin
  Create;
  FStrings.AddStrings(aStrings)
end;

function TStringsEnumerator.Count: Integer;
begin
  Result:= FStrings.Count
end;

{ TUnkEnumerator }

constructor TUnkEnumerator.Create;
begin
  inherited Create;
  FList:= TInterfaceList.Create
end;

constructor TUnkEnumerator.CreateGroupEnumerator(GroupList: TGroupList;
  Scope: OPCENUMSCOPE);
var
  i: Integer;
  G: TGroupImpl;
begin
  Create;
  for i:= 0 to GroupList.Count - 1 do
  begin
    G:= GroupList.Group(i);
    if G.InScope(Scope) then
      FList.Add(G)
  end
end;

constructor TUnkEnumerator.CreateClone(aList: TInterfaceList);
var
  i: Integer;
begin
  Create;
  for i:= 0 to aList.Count - 1 do
    FList.Add(aList[i])
end;

destructor TUnkEnumerator.Destroy;
begin
  FList.Free;
  inherited Destroy
end;

function TUnkEnumerator.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
  i: Integer;
  ItemCount: Integer;
begin
  ItemCount:= FList.Count - FIndex;
  if celt > ItemCount then
  begin
    Result:= S_FALSE
  end else
  begin
    ItemCount:= celt;
    Result:= S_OK
  end;
  if Assigned(pceltFetched) then
    pceltFetched^:= ItemCount;
  if ItemCount > 0 then
  begin
    i:= 0;
    while i < ItemCount do
    begin
      TUnknownList(elt)[i]:= FList[FIndex];
      Inc(FIndex);
      Inc(i)
    end
  end
end;

function TUnkEnumerator.Skip(celt: Longint): HResult;
begin
  if (FIndex + celt) <= FList.Count then
  begin
    Inc(FIndex, celt);
    Result:= S_OK
  end else
  begin
    FIndex:= FList.Count;
    Result:= S_FALSE
  end
end;

function TUnkEnumerator.Reset: HResult;
begin
  FIndex:= 0;
  Result:= S_OK;
end;

function TUnkEnumerator.Clone(out enm: IEnumUnknown): HResult;
begin
  try
    enm:= TUnkEnumerator.CreateClone(FList);
    Result:= S_OK
  except
    Result:= E_UNEXPECTED
  end;
end;

function TUnkEnumerator.Count: Integer;
begin
  Result:= FList.Count
end;

{ TClientInfo }

function TClientInfo.ShutdownRequest(const Reason: String): Boolean;
var
  WideReason: WideString;
begin
  Result:= Assigned(FOpcShutdown);
  if Result then
  begin
    WideReason:= Reason;
    try
      FOPCShutdown.ShutdownRequest(PWideChar(WideReason))
    except
      on E: EOleSysError do
        GOpcItemServer.ClientCallbackError(E, Self, ccShutdownRequest)
    end
  end
end;

function TClientInfo.ShutdownRequestAvail: Boolean;
begin
  Result:= Assigned(FOpcShutdown)
end;

{ TServerImpl }

function TServerImpl.AddGroup(szName: POleStr; bActive: BOOL;
  dwRequestedUpdateRate: DWORD; hClientGroup: OPCHANDLE;
  pTimeBias: PLongint; pPercentDeadband: PSingle; dwLCID: DWORD;
  out phServerGroup: OPCHANDLE; out pRevisedUpdateRate: DWORD;
  const riid: TIID; out ppUnk: IUnknown): HResult;
var
  NewGroup: TGroupImpl;
  Tzi: TTimeZoneInformation;
  DB: Single;
  i: Integer;
begin
  ppUnk:= nil;
  NewGroup:= nil;
  try
    if Assigned(pTimeBias) then
      Tzi.Bias:= pTimeBias^
    else
      GetTimeZoneInformation(Tzi);     {cf 1.14.8}
   if Assigned(pPercentDeadband) then  {cf 1.01.3}
     DB:= pPercentDeadband^
   else
     DB:= 0;
    NewGroup:= TGroupImpl.Create(Self, szName, bActive, dwRequestedUpdateRate, hClientGroup,
      Tzi.Bias, DB, dwLCID);
    pRevisedUpdateRate:= NewGroup.UpdateRate;
    phServerGroup:= OPCHANDLE(NewGroup);
    if not NewGroup.GetInterface(riid, ppUnk) then
      raise EOpcError.Create(E_NOINTERFACE);
    if (pRevisedUpdateRate = dwRequestedUpdateRate) then {cf 1.01.4}
      Result:= S_OK
    else
      Result:= OPC_S_UNSUPPORTEDRATE
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
      if Assigned(NewGroup) then
      begin  {cf 1.13.4}
        i:= FGroupList.IndexOf(NewGroup.Name);
        if i <> -1 then
          FGroupList.Delete(i);
        NewGroup.Free
      end
    end
  end
end;

function TServerImpl.BrowseAccessPaths(szItemID: POleStr;
  out ppIEnumString: IEnumString): HResult;
begin
  ppIEnumString:= nil;
  Result:= E_NOTIMPL
end;

function TServerImpl.BrowseOPCItemIDs(dwBrowseFilterType: OPCBROWSETYPE;
  szFilterCriteria: POleStr; vtDataTypeFilter: TVarType;
  dwAccessRightsFilter: DWORD; out ppIEnumString: IEnumString): HResult;
var
  se: TStringsEnumerator;
begin
  ppIEnumString:= nil;
  try
    TestParamRange(dwBrowseFilterType, OPC_BRANCH, OPC_FLAT);
    TestParamRange(dwAccessRightsFilter, 0, OPC_READABLE or OPC_WRITABLE);
    GOpcItemServer.InitBrowsing(FBrowsePos);
    se:= TStringsEnumerator.CreateOpcItemsEnumerator(GOpcItemServer,
      FBrowsePos, dwBrowseFilterType, szFilterCriteria,
      vtDataTypeFilter, dwAccessRightsFilter);
    ppIEnumString:= se;
    if se.Count = 0 then  {cf 1.13.5}
      Result:= S_FALSE
    else
      Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TServerImpl.ChangeBrowsePosition(
  dwBrowseDirection: OPCBROWSEDIRECTION; szString: POleStr): HResult;

function NullString: Boolean; {cf 1.15.2}
begin
  Result:= not Assigned(szString) or (szString^ = #0)
end;

procedure GoToNode(Start: TItemIdList);
var
  Node: TNamespaceNode;
begin
  Node:= Start.Find(szString);
  if not Assigned(Node) or not (Node is TItemIdList) then
    raise EOpcError.Create(E_INVALIDARG);
  FBrowsePos:= TItemIdList(Node)
end;

begin
  if not GOpcItemServer.HierarchicalBrowsing then
  begin
    Result:= E_FAIL
  end else
  begin
    try
      TestParamRange(dwBrowseDirection, OPC_BROWSE_UP, OPC_BROWSE_TO);
      with GOpcItemServer do
      begin
        InitBrowsing(FBrowsePos);
        LockServerList;
        try
          Result := E_FAIL;
          case dwBrowseDirection of
            OPC_BROWSE_UP:
            begin
              if FBrowsePos = FRootNode then
               Exit;
              FBrowsePos:= FBrowsePos.Parent
            end;
            OPC_BROWSE_DOWN:
            begin
              {szString can be more than one branch ie AREA1.REACTOR10.
              I am not sure if this is strictly allowed but it seems sensible}
              if NullString then   {cf 1.15.2}
                raise EOpcError.Create(E_INVALIDARG);
              GoToNode(FBrowsePos)
            end;
            OPC_BROWSE_TO:
            begin
              if NullString then   {cf 1.15.2}
                FBrowsePos:= FRootNode
              else
                GoToNode(FRootNode)
            end
          end;
        finally
          UnlockServerList
        end
      end;
      Result:= S_OK
    except
      on E: EOpcError do
        Result:= E.ErrorCode
    end
  end
end;

function TServerImpl.CreateGroupEnumerator(dwScope: OPCENUMSCOPE;
  const riid: TIID; out ppUnk: IUnknown): HResult;
var
  StringsEnum: TStringsEnumerator;
  UnkEnum: TUnkEnumerator;
begin
  ppUnk:= nil;
  try
    if IsEqualGUID(riid, IEnumUnknown) then
    begin
      UnkEnum:= TUnkEnumerator.CreateGroupEnumerator(FGroupList, dwScope);
      if UnkEnum.Count = 0 then
        Result:= S_FALSE
      else
        Result:= S_OK;
      ppUnk:= UnkEnum
    end else
    if IsEqualGUID(riid, IEnumString) then
    begin
      StringsEnum:= TStringsEnumerator.CreateGroupEnumerator(FGroupList, dwScope);
      if StringsEnum.Count = 0 then
        Result:= S_FALSE
      else
        Result:= S_OK;
      ppUnk:= StringsEnum
    end else
    begin
      Result:= E_NOINTERFACE
    end
  except {in case iterator constructors fail}
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

destructor TServerImpl.Destroy;
begin
{$IFDEF GLD}
  SendDebug(Format('GLD: ServerDestroy GroupCount = %d', [GroupCount]));
{$ENDIF}
  {notify in main thread}
  with LockServerList do
  try
    Remove(Self)
  finally
    UnlockServerList
  end;
  if FServerState = OPC_STATUS_RUNNING then
    SendMessage(OpcWindow, CM_NOTIFICATION, nfClientDisconnect, Integer(Self));
  FGroupList.Free; {what about dangling references?}
  FConnectionPoint.Free;
  inherited Destroy       {moved from start of destructor}
end;

function TServerImpl.EnumConnectionPoints(
  out enumconn: IEnumConnectionPoints): HResult;
begin
  enumconn:= TEnumerateCP.Create(FConnectionPoint, 0); {cf 1.01.5}
  Result:= S_OK
end;

function TServerImpl.FindConnectionPoint(const iid: TIID;
  out cp: IConnectionPoint): HResult;
begin
  if IsEqualGUID(FConnectionPoint.FIID, iid) then
  begin
    cp:= FConnectionPoint;
    Result:= S_OK
  end else
  begin
    Result:= CONNECT_E_NOCONNECTION
  end
end;

function TServerImpl.FillProperties(
            szItemID: POleStr;
            bReturnPropertyValues:      BOOL;
            dwPropertyCount:            DWORD;
            pdwPropertyIDs:             PDWORDARRAY;
      var   ItemProperties:             OPCITEMPROPERTIES): HResult;
var
  I: Integer;
  ItemResult: TServerItemRef;
  ValueRead: Boolean;
  ItemQuality: Word;
  SysTime: TSystemTime;
  ItemTimestamp: TFiletime;
  ItemValue: OleVariant;
  OleTime: TOleDate;
  IsError: Boolean;
  PropertyID: LongWord;
  ListProperties: Boolean;
  ItemProperty: IItemProperty;

procedure DoReadValue;
begin
  if not ValueRead then
  begin
    ItemQuality:= OPC_QUALITY_GOOD;
    ItemValue:= GOpcItemServer.GetItemVQT(ItemResult.ItemHandle, ItemQuality, ItemTimestamp);
    ValueRead:= true
  end
end;

begin
  try
    IsError := False;
    if dwPropertyCount = 0 then
      ListProperties := True
    else
      ListProperties := False;

    try
      ItemResult:= TServerItemRef.GetItem(szItemID);
      try
        ValueRead:= false;
        if ListProperties then
        begin
          dwPropertyCount := StdItemPropCount;
          if Assigned(ItemResult.ItemProperties) then
            dwPropertyCount := dwPropertyCount + Cardinal(ItemResult.ItemProperties.Count);
        end;

        ItemProperties.hrErrorID := S_OK;
        ItemProperties.dwNumProperties := dwPropertyCount;
        ItemProperties.pItemProperties := ZeroAllocation(dwPropertyCount*SizeOf(OPCITEMPROPERTY));
        for I := 0 to dwPropertyCount-1 do
        begin
          if not ListProperties then
            PropertyID := pdwPropertyIDs^[I]
          else
            begin
              if I < StdItemPropCount then
                PropertyID := StdItemPropID[I]
              else
                PropertyID := ItemResult.ItemProperties.GetPropertyItem(I - StdItemPropCount).Pid
            end;
          ItemProperties.pItemProperties^[I].hrErrorID := S_OK;
          ItemProperties.pItemProperties^[I].dwPropertyID := PropertyID;
          ItemProperties.pItemProperties^[I].szItemID := nil;
          case PropertyID of
            OPC_PROPERTY_DATATYPE:
            begin
              ItemProperties.pItemProperties^[I].vtDataType := VT_I2;
              ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(OPC_PROPERTY_DESC_DATATYPE);
              if bReturnPropertyValues then
               begin
                 ItemProperties.pItemProperties^[I].vValue := ItemResult.CanonicalDataType;
                 VariantChangeType(ItemProperties.pItemProperties^[I].vValue, ItemProperties.pItemProperties^[I].vValue, 0, VT_I2);
               end;
            end;
            OPC_PROPERTY_VALUE:
            begin
              DoReadValue;
              ItemProperties.pItemProperties^[I].vtDataType := VT_VARIANT;
              ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(OPC_PROPERTY_DESC_VALUE);
              if bReturnPropertyValues then
               ItemProperties.pItemProperties^[I].vValue := ItemValue
            end;
            OPC_PROPERTY_QUALITY:
            begin
              ItemProperties.pItemProperties^[I].vtDataType := VT_I2;
              ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(OPC_PROPERTY_DESC_QUALITY);
              if bReturnPropertyValues then
               begin
                 DoReadValue;
                 ItemProperties.pItemProperties^[I].vValue := ItemQuality;
                 VariantChangeType(ItemProperties.pItemProperties^[I].vValue, ItemProperties.pItemProperties^[I].vValue, 0, VT_I2);
               end;
            end;
            OPC_PROPERTY_TIMESTAMP:
            begin
              ItemProperties.pItemProperties^[I].vtDataType := VT_DATE;
              ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(OPC_PROPERTY_DESC_TIMESTAMP);
              if bReturnPropertyValues then
               begin
                 DoReadValue;
                 FileTimeToSystemTime(ItemTimestamp, SysTime);
                 SystemTimeToVariantTime(SysTime, OleTime);
                 ItemProperties.pItemProperties^[I].vValue := OleTime;
                 VariantChangeType(ItemProperties.pItemProperties^[I].vValue, ItemProperties.pItemProperties^[I].vValue, 0, VT_DATE);
               end;
            end;
            OPC_PROPERTY_ACCESS_RIGHTS:
            begin
              ItemProperties.pItemProperties^[I].vtDataType := VT_I4;
              ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(OPC_PROPERTY_DESC_ACCESS_RIGHTS);
              if bReturnPropertyValues then
               ItemProperties.pItemProperties^[I].vValue := Integer(ItemResult.AccessRights);
            end;
            OPC_PROPERTY_SCAN_RATE:
            begin
              ItemProperties.pItemProperties^[I].vtDataType := VT_R4;
              ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(OPC_PROPERTY_DESC_SCAN_RATE);
              if bReturnPropertyValues then
               begin
                 ItemProperties.pItemProperties^[I].vValue := GOpcItemServer.MaxUpdateRate;
                 VariantChangeType(ItemProperties.pItemProperties^[I].vValue, ItemProperties.pItemProperties^[I].vValue, 0, VT_R4);
               end;
            end;
          else
            begin
              ItemProperty := nil;
              if Assigned(ItemResult.ItemProperties) then
                ItemProperty := ItemResult.ItemProperties.GetProperty(PropertyID);
              if ItemProperty <> nil then
              begin
                ItemProperties.pItemProperties^[I].vtDataType := ItemProperty.DataType;
                ItemProperties.pItemProperties^[I].szDescription := StringToLPOLESTR(ItemProperty.Description);
                if bReturnPropertyValues then
                begin
                  ItemProperties.pItemProperties^[I].vValue := ItemProperty.GetPropertyValue;
                  VariantChangeType(ItemProperties.pItemProperties^[I].vValue, ItemProperties.pItemProperties^[I].vValue, 0, ItemProperty.DataType);
                end;
              end else
              begin
                ItemProperties.hrErrorID := S_FALSE;
                ItemProperties.pItemProperties^[I].hrErrorID := OPC_E_INVALID_PID;
                IsError := True;
              end;
            end;
          end
        end;
      finally
        ItemResult.ReleaseNonGroupReference;
      end;
    except
      on E: EOpcError do
      begin
        ItemProperties.hrErrorID := E.ErrorCode;
        IsError := True;
      end;
    end;

    if not IsError then
      Result := S_OK
    else
      Result := S_FALSE;
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
    end
  end;
end;

function TServerImpl.GetProperties(
            dwItemCount:                DWORD;
            pszItemIDs:                 POleStrList;
            bReturnPropertyValues:      BOOL;
            dwPropertyCount:            DWORD;
            pdwPropertyIDs:             PDWORDARRAY;
      out   ppItemProperties:           POPCITEMPROPERTIESARRAY): HResult;
var
  I: Integer;
  IsError: Boolean;
begin
  try
    ppItemProperties := nil;
    IsError := False;

    CheckCount(dwItemCount, pszItemIDs);

    ppItemProperties := ZeroAllocation(dwItemCount*SizeOf(OPCITEMPROPERTIES));
    for I := 0 to dwItemCount-1 do
    begin
      if FillProperties(pszItemIDs[I], bReturnPropertyValues, dwPropertyCount, pdwPropertyIDs, ppItemProperties^[I]) <> S_OK then
        IsError := True;
    end;

    if not IsError then
      Result := S_OK
    else
      Result := S_FALSE;
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
    end
  end;
end;

function TServerImpl.Browse(
            szItemID:                   POleStr;
      var   pszContinuationPoint:       POleStr;
            dwMaxElementsReturned:      DWORD;
            dwBrowseFilter:             OPCBROWSEFILTER;
            szElementNameFilter:        POleStr;
            szVendorFilter:             POleStr;
            bReturnAllProperties:       BOOL;
            bReturnPropertyValues:      BOOL;
            dwPropertyCount:            DWORD;
            pdwPropertyIDs:             PDWORDARRAY;
      out   pbMoreElements:             BOOL;
      out   pdwCount:                   DWORD;
      out   ppBrowseElements:           POPCBROWSEELEMENTARRAY): HResult; stdcall;
var
  I : Integer;
  StartI, Index : LongWord;
  List : TItemIdList;
  NSNode : TNamespaceNode;
function FilterNameSpaceNode(NameSpaceNode: TNamespaceNode; ElementNameFilter: string; BrowseFilter: DWORD) : Boolean;
begin
  Result := False;
  if (ElementNameFilter = '') or
     GOpcItemServer.FilterFunction(ElementNameFilter, [], 0, NameSpaceNode.Name, [], 0) then
  begin
    if (BrowseFilter in [OPC_BROWSE_FILTER_ALL, OPC_BROWSE_FILTER_BRANCHES]) and
       (NameSpaceNode is TItemIdList) or
       (BrowseFilter in [OPC_BROWSE_FILTER_ALL, OPC_BROWSE_FILTER_ITEMS]) and
       (NameSpaceNode is TNamespaceItem) then
      Result := True;
  end;
end;

begin
  Result := S_OK;

  pbMoreElements := False;
  List := nil;
  try
    TestParamRange(dwBrowseFilter, OPC_BROWSE_FILTER_ALL, OPC_BROWSE_FILTER_ITEMS);
    try
      if pszContinuationPoint^ <> #0 then
      begin
        StartI := StrToInt(pszContinuationPoint);
        FreeAndNull(pszContinuationPoint);
        pszContinuationPoint := StringToLPOLESTR('');
      end else
        StartI := 0;
    except
      raise EOpcError.Create(OPC_E_INVALIDCONTINUATIONPOINT);
    end;

    if szItemID^ = #0 then
     List := GOpcItemServer.FRootNode
    else
     begin
       NSNode := GOpcItemServer.FRootNode.Find(szItemID);
       if NSNode = nil then
         raise EOpcError.Create(OPC_E_UNKNOWNITEMID);
       if NSNode is TItemIdList then
        List := TItemIdList(NSNode);
     end;

    pdwCount := 0;
    for I := StartI to List.ChildCount - 1 do
      if FilterNameSpaceNode(List.Child(I), szElementNameFilter, dwBrowseFilter) then
        Inc(pdwCount);

    if (dwMaxElementsReturned > 0) and (pdwCount > dwMaxElementsReturned) then
      pdwCount := dwMaxElementsReturned;

    ppBrowseElements := ZeroAllocation(pdwCount*SizeOf(OPCBROWSEELEMENT));

    Index := 0;
    for I := StartI to List.ChildCount - 1 do
    begin
      if FilterNameSpaceNode(List.Child(I), szElementNameFilter, dwBrowseFilter) then
      begin
        if Index = pdwCount then
        begin
          FreeAndNull(pszContinuationPoint);
          pszContinuationPoint := StringToLPOLESTR(IntToStr(I));
          Break;
        end;

        ppBrowseElements[Index].szName := StringToLPOLESTR(List.Child(I).Name);
        ppBrowseElements[Index].szItemID := StringToLPOLESTR(List.Child(I).Path);
        if List.Child(I) is TItemIDList then
        begin
          ppBrowseElements[Index].dwFlagValue := OPC_BROWSE_HASCHILDREN;
          ppBrowseElements[Index].ItemProperties.hrErrorID := OPC_E_UNKNOWNITEMID;
        end else
        begin
          ppBrowseElements[Index].dwFlagValue := OPC_BROWSE_ISITEM;
          if bReturnAllProperties or (dwPropertyCount > 0) then
            FillProperties(ppBrowseElements[Index].szItemID, bReturnPropertyValues, dwPropertyCount, pdwPropertyIDs, ppBrowseElements[Index].ItemProperties);
        end;
        Inc(Index);
      end;
    end;
  except
    on E: EOPCError do
      Result := E.ErrorCode;
  end;
end;

function TServerImpl.Read(
            dwCount:                    DWORD;
            pszItemIDs:                 POleStrList;
            pdwMaxAge:                  PDWORDARRAY;
      out   ppvValues:                  POleVariantArray;
      out   ppwQualities:               PWordArray;
      out   ppftTimeStamps:             PFileTimeArray;
      out   ppErrors:                   PResultList): HResult;
var
  I : Integer;
  IsError : Boolean;
  ItemResult: TServerItemRef;
  ItemQuality: Word;
  ItemTimestamp: TFiletime;
  ItemValue: OleVariant;
  ActualTimestamp: TFileTime;
begin
  IsError := False;
  try
    CheckCount(dwCount, pszItemIDs);
    ppvValues := ZeroAllocation(dwCount * SizeOf(OleVariant));
    ppwQualities := ZeroAllocation(dwCount * SizeOf(Word));
    ppftTimeStamps := ZeroAllocation(dwCount * SizeOf(TFileTime));
    ppErrors := ZeroAllocation(dwCount * SizeOf(HRESULT));
    GOpcItemServer.GetTimestamp(ActualTimestamp);
    for I := 0 to dwCount - 1 do
    try
      ItemResult:= TServerItemRef.GetItem(pszItemIDs[I]);
      try
        ItemQuality:= OPC_QUALITY_GOOD;
        ItemResult.GetMaxAgeVQT(pdwMaxAge[I], ActualTimestamp, ItemValue, ItemQuality, ItemTimestamp);
        ppvValues[I]:= ItemValue;
        ppwQualities[I]:= ItemQuality;
        ppftTimeStamps[I]:= ItemTimestamp;
      finally
        ItemResult.ReleaseNonGroupReference
      end;
    except
      on E: EOpcError do
      begin
        ppErrors[I]:= E.ErrorCode;
        IsError:= True;
      end;
    end;

    if not IsError then
      Result:= S_OK
    else
      Result:= S_FALSE;
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end;
end;

function TServerImpl.WriteVQT(
            dwCount:                    DWORD;
            pszItemIDs:                 POleStrList;
            pItemVQT:                   POPCITEMVQTARRAY;
      out   ppErrors:                   PResultList): HResult;
var
  IsError : Boolean;
  I : Integer;
  ItemResult: TServerItemRef;
begin
  IsError := False;
  try
    CheckCount(dwCount, pszItemIDs);
    ppErrors := ZeroAllocation(dwCount * SizeOf(HRESULT));
    for I := 0 to dwCount - 1 do
    try
      ItemResult:= TServerItemRef.GetItem(pszItemIDs[I]);
      try
        if VarType(pItemVQT[I].vDataValue) = VT_EMPTY then
          raise EOpcError.Create(OPC_E_BADTYPE);

        ItemResult.SetItemVQT(pItemVQT[I]);
      finally
        ItemResult.ReleaseNonGroupReference
      end;
    except
      on E: EOPCError do
      begin
        ppErrors[I] := E.ErrorCode;
        IsError := True;
      end;
    end;

    if not IsError then
      Result := S_OK
    else
      Result := S_FALSE;
  except
    on E: EOpcError do
      Result := E.ErrorCode;
  end;
end;

function TServerImpl.ClearBrowsingContext: String;
begin
  if Assigned(FBrowsePos) then
  begin
    Result:= FBrowsePos.Path;
    FBrowsePos:= nil
  end else
  begin
    Result:= ''
  end
end;

function TServerImpl.GetGroupByName(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult;
begin
  ppUnk:= nil;
  try
    FGroupList.GroupByName(szName, riid, ppUnk);
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

{calling SysAllocString causes that odd
'break in ntdll.dll problem. V. strange}

function MySysAllocString(const S: WideString): POleStr;
var
  BufSize: Integer;
begin
  BufSize:= (Length(S) + 1) * SizeOf(WideChar);
  Result:= CoTaskMemAlloc(BufSize);
  if Result = nil then
    raise EOpcError.Create(E_OUTOFMEMORY);
  if S = '' then
    Result^:= #0
  else
    Move(Pointer(S)^, Result^, BufSize)
end;

{cf 1.14.7}
function TServerImpl.GetItemID(szItemDataID: POleStr;
  out szItemID: POleStr): HResult;
var
  Node: TNamespaceNode;
  ItemDataId: String;
begin
  szItemID:= nil;
  try
    Result:= S_OK;
    if szItemDataId <> nil then
    begin
      if szItemDataId^ = #0 then
      begin
        szItemID:= CoTaskMemAlloc(SizeOf(WideChar));
        szItemID^:= #0
      end else
      begin
        with GOpcItemServer do
        begin
          InitBrowsing(FBrowsePos);
          LockServerList;
          try
            if GOpcItemServer.HierarchicalBrowsing then
            begin
              ItemDataId:= szItemDataId;
              Node:= FBrowsePos.Find(ItemDataId);
              {it might be a fully qualified path. Try the root node}
              if not Assigned(Node) then
                Node:= FRootNode.Find(ItemDataId);
              if not Assigned(Node) then
                raise EOpcError.Create(E_INVALIDARG);
              szItemId:= MySysAllocString(Node.Path)
            end else
            begin
              if FRootNode.Find(szItemDataId) = nil then
                raise EOpcError.Create(E_INVALIDARG);
              szItemID:= MySysAllocString(szItemDataID);
              Result:= S_OK
            end
          finally
            UnlockServerList
          end;

{ use permanent browse listing instead
          ItemResult:= TServerItemRef.GetItem(szItemDataID);
          try
            szItemID:= MySysAllocString(szItemDataID);
            if szItemID = nil then
              Result:= E_OUTOFMEMORY
            else
              Result:= S_OK
          finally
            ItemResult.ReleaseNonGroupReference
          end
}
        end
      end
    end
  except
    on E: EOpcError do
    begin
      FreeAndNull(szItemID);
      Result:= E.ErrorCode
    end
  end
end;

function TServerImpl.GetItemProperties(szItemID: POleStr; dwCount: DWORD;
  pdwPropertyIDs: PDWORDARRAY; out ppvData: POleVariantArray;
  out ppErrors: PResultList): HResult;
var
  ItemResult: TServerItemRef;
  i: Integer;
  IsError: Boolean;
  ItemProperties: OPCITEMPROPERTIES;
begin
  ppvData:= nil;
  ppErrors:= nil;
  try
    CheckCount(dwCount, pdwPropertyIDs);
    ItemResult:= TServerItemRef.GetItem(szItemID);
    try
      ppvData:= ZeroAllocation(dwCount*SizeOf(OleVariant)); {cf 1.12.3}
      ppErrors:= CheckAllocation(dwCount*SizeOf(HRESULT));
      IsError:= False;
      if FillProperties(szItemID, True, dwCount, pdwPropertyIDs, ItemProperties) <> S_OK then
        IsError := True;
      try
        for i:= 0 to dwCount - 1 do
        begin
          ppvData^[i]:=ItemProperties.pItemProperties[i].vValue;
          ppErrors^[i]:=ItemProperties.pItemProperties[i].hrErrorID;
        end;
      finally
        FreeOPCItemProperties(ItemProperties);
      end;
    finally
      ItemResult.ReleaseNonGroupReference
    end;
    if not IsError then
      Result:= S_OK
    else
      Result:= S_FALSE
  except
    on E: EOpcError do
    begin
      FreeAndNull(ppvData);
      FreeAndNull(ppErrors);
      Result:= E.ErrorCode;
    end
  end;
end;

function TServerImpl.GetLocaleID(out pdwLcid: TLCID): HResult;
begin
  pdwLcid:= FLCID;
  Result:= S_OK
end;

function TServerImpl.GetStatus(
  out ppServerStatus: POPCSERVERSTATUS): HResult;
var
  VI: Pointer;
  VersionSize, Z: DWORD;
  FFI: PVSFixedFileInfo;
  ItemLength: DWORD;
  Fn: String;

begin
  ppServerStatus:= nil;
  try
    ppServerStatus:= CheckAllocation(SizeOf(OPCSERVERSTATUS));
    FillChar(ppServerStatus^, SizeOf(OPCSERVERSTATUS), 0);
    with ppServerStatus^ do
    begin
      ftStartTime:= OpcItemServer.StartupTime;
      GOpcItemServer.GetTimestamp(ftCurrentTime); {cf 1.14.28}
      ftLastUpdateTime:= FLastUpdateTime;
      dwServerState:= FServerState;
      dwGroupCount:= FGroupList.Count;
      dwBandWidth:= DWORD(-1);  {na}
      Fn:= ParamStr(0);
      VersionSize:= GetFileVersionInfoSize(PChar(Fn), Z);
      if VersionSize > 0 then
      begin
        GetMem(VI, VersionSize);
        try
          if GetFileVersionInfo(PChar(Fn), 0, VersionSize, VI) then
          begin
            VerQueryValue(VI, '\', Pointer(FFI), ItemLength);
            with FFI^ do
            begin
              wMajorVersion:= HIWORD(dwFileVersionMS);
              wMinorVersion:= LOWORD(dwFileVersionMS);
              wBuildNumber:= LOWORD(dwFileVersionLS)
            end;
          end
        finally
          FreeMem(VI)
        end;
      end;
      szVendorInfo:= StringToLPOLESTR(GOpcItemServer.FVendorInfo)
    end;
    Result:= S_OK
  except
    on EOutOfMemory do
    begin
      FreeAndNull(ppServerStatus);
      Result:= E_OUTOFMEMORY
    end;
    on E: EOpcError do
    begin
      FreeAndNull(ppServerStatus);
      Result:= E.ErrorCode
    end
  end
end;

function TServerImpl.Group(i: Integer): TGroupInfo;
begin
  Result:= FGroupList.Group(i)
end;

function TServerImpl.GroupCount: Integer;
begin
  Result:= FGroupList.Count
end;

procedure TServerImpl.Initialize;
begin
  inherited Initialize;
  {if this exception is raised then someone is trying to
   create an OPCServer locally. This should be impossible}
  if not Assigned(GOpcItemServer) then
    raise EOpcServer.CreateRes(@SNoItemServer);
{$IFDEF Evaluation}
  if OpcItemServer.ClientCount >= MaxEvalClients then
    raise EOleSysError.Create('Too many clients', CLASS_E_NOTLICENSED, 0);
{$ENDIF}
  FConnectionPoint:= TConnectionPoint.Create(Self, IOPCShutdown, SinkConnect);
  FGroupList:= TGroupList.Create;
  FLCID:= LOCALE_SYSTEM_DEFAULT; {cf 1.01.1 was GetUserDefaultLCID}
  with LockServerList do
  try
    Add(Self)
  finally
    UnlockServerList
  end;
  {notify in main thread}
  FServerState:= OPC_STATUS_RUNNING;
  SendMessage(OpcWindow, CM_NOTIFICATION, nfClientConnect, Integer(Self))
end;

function TServerImpl.LookupItemIDs(szItemID: POleStr; dwCount: DWORD;
  pdwPropertyIDs: PDWORDARRAY; out ppszNewItemIDs: POleStrList;
  out ppErrors: PResultList): HResult;

procedure FreeAndNullNewItemIDs;
var
  i: Integer;
begin
  if Assigned(ppszNewItemIDs) then
  for i:= 0 to dwCount - 1 do
  begin
    if Assigned(ppszNewItemIDs^[i]) then
      CoTaskMemFree(ppszNewItemIDs^[i])
  end
end;

var
  ItemResult: TServerItemRef;
  i: Integer;
{I just fail all these (individually)}
begin
  ppszNewItemIDs:= nil;
  ppErrors:= nil;
  try
    CheckCount(dwCount, pdwPropertyIDs);
    ItemResult:= TServerItemRef.GetItem(szItemID);
    try
      ppszNewItemIDs:= CheckAllocation(dwCount*SizeOf(PWideChar));
      FillChar(ppszNewItemIDs^, dwCount*SizeOf(PWideChar), 0);
      ppErrors:= CheckAllocation(dwCount*SizeOf(HRESULT));
      for i:= 0 to dwCount - 1 do
      begin
        ppszNewItemIDs^[i]:= nil; {StringToLPOLESTR('');}
        if pdwPropertyIDs^[i] <= 6 then
         ppErrors^[i] := OPC_E_INVALID_PID
        else
         ppErrors^[i]:= E_FAIL;
      end
    finally
      ItemResult.ReleaseNonGroupReference
    end;
    Result:= S_FALSE  {cf 13.1.6}
  except
    on E: EOpcError do
    begin
      FreeAndNull(ppErrors);
      FreeAndNullNewItemIDs;
      Result:= E.ErrorCode
    end
  end
end;


function TServerImpl.OPCCommonGetErrorString(dwError: HResult;
  out ppString: POleStr): HResult;
begin
  Result:= OPCServerGetErrorString(dwError, FLCID, ppString)
end;

function TServerImpl.OpcServerGetErrorString(dwError: HResult;
  dwLocale: TLCID; out ppString: POleStr): HResult;
var
  Res: String;
begin
  try
    ppString:= nil;
    Res:= OpcItemServer.GetErrorString(dwError, dwLocale);
    ppString:= StringToLPOLESTR(Res);
    Result:= S_OK;
  except
    on E:EOpcError do
      Result:= E.ErrorCode
  end
end;

function TServerImpl.QueryAvailableLocaleIDs(out pdwCount: UINT;
  out pdwLcid: PLCIDARRAY): HResult;
{cf 1.13.2}
begin
  pdwLcid:= nil;
  try
    pdwCount:= 3;
    pdwLcid:= CheckAllocation(SizeOf(LCID) * pdwCount);
    pdwLcid^[0]:= LOCALE_SYSTEM_DEFAULT;
    pdwLcid^[1]:= GetSystemDefaultLCID;
    pdwLcid^[2]:= GetUserDefaultLCID;
    Result:= S_OK
  except
    on E: EOpcError do
    begin
      FreeAndNull(pdwLcid);
      Result:= E.ErrorCode
    end
  end
end;

function TServerImpl.QueryAvailableProperties(szItemID: POleStr;
  out pdwCount: DWORD; out ppPropertyIDs: PDWORDARRAY;
  out ppDescriptions: POleStrList;
  out ppvtDataTypes: PVarTypeList): HResult;
var
  ItemResult: TServerItemRef;
  i: Integer;
  ItemProperties: OPCITEMPROPERTIES;
begin
  ppPropertyIDs:= nil;
  ppDescriptions:= nil;
  ppvtDataTypes:= nil;
  try
    ItemResult:= TServerItemRef.GetItem(szItemID);
    try
      FillProperties(szItemID, False, 0, nil, ItemProperties);

      try
        pdwCount := ItemProperties.dwNumProperties;
        ppPropertyIDs:= CheckAllocation(pdwCount*SizeOf(DWORD));
        ppDescriptions:= CheckAllocation(pdwCount*SizeOf(PWideChar));
        ppvtDataTypes:= CheckAllocation(pdwCount*SizeOf(TVarType));
        for i:= 0 to pdwCount - 1 do
        begin
          ppPropertyIDs^[i]:= ItemProperties.pItemProperties^[i].dwPropertyID;
          ppDescriptions^[i]:= StringToLPOLESTR(ItemProperties.pItemProperties^[i].szDescription);
          ppvtDataTypes^[i]:= ItemProperties.pItemProperties^[i].vtDataType;
        end;
      finally
        FreeOPCItemProperties(ItemProperties);
      end;
      ppvtDataTypes^[OPC_PROPERTY_DATATYPE]:= ItemResult.CanonicalDataType;
    finally
      ItemResult.ReleaseNonGroupReference
    end;
    Result:= S_OK
  except
    on E: EOpcError do
    begin
      FreeAndNull(ppPropertyIDs);
      FreeAndNull(ppDescriptions);
      FreeAndNull(ppvtDataTypes);
      Result:= E.ErrorCode
    end
  end
end;

function TServerImpl.QueryOrganization(
  out pNameSpaceType: OPCNAMESPACETYPE): HResult;
begin
  Result:= S_OK;
  if GOpcItemServer.HierarchicalBrowsing then
    pNameSpaceType:= OPC_NS_HIERARCHIAL
  else
    pNameSpaceType:= OPC_NS_FLAT
end;

function TServerImpl.RemoveGroup(hServerGroup: OPCHANDLE;
  bForce: BOOL): HResult;
var
  Group: TGroupImpl absolute hServerGroup;
  i: Integer;
begin
  i:= FGroupList.IndexOfObject(Group);
  if i = -1 then
  begin
    Result:= E_INVALIDARG  {was E_OPC_INVALIDHANDLE cf 1.13.3}
  end else
  begin
    {$IFDEF GLD}
    with Group do
      SendDebug(Format('GLD: RemoveGroup %s RC = %d', [Name, ReferenceCount]));
    {$ENDIF}
    FGroupList.Delete(i);
    if (Group.ReferenceCount = 0) or bForce then
    begin
      if Group.ReferenceCount = 0 then
       Group.Free
      else
       Group.FDeleted:= true;
      Result:= S_OK
    end else
    begin
      Group.FDeleted:= true;
      Result:= OPC_S_INUSE;
    end
  end
end;

procedure TServerImpl.RestoreBrowsingContext(const S: String);
var
  Root: TItemIdList;
  Node: TNamespaceNode;
begin
  Root:= GOpcItemServer.FRootNode;
  if Assigned(Root) then
  begin
    Node:= Root.Find(S);
    if Assigned(Node) and (Node is TItemIdList) then
      FBrowsePos:= TItemIdList(Node)
    else
      FBrowsePos:= nil
  end else
  begin
    FBrowsePos:= nil
  end
end;

function TServerImpl.SetClientName(szName: POleStr): HResult;
begin
  FClientName:= szName;
  Result:= S_OK;
  {notify in main thread}
  SendMessage(OpcWindow, CM_NOTIFICATION, nfClientSetName, Integer(Self))
end;

function TServerImpl.SetLastUpdate: TFileTime;
begin
  GOpcItemServer.GetTimestamp(FLastUpdateTime); {cf 1.14.28}
  Result:= FLastUpdateTime
end;

function TServerImpl.SetLocaleID(dwLcid: TLCID): HResult;
{cf 1.13.2}
begin
  if (dwLCID = LOCALE_SYSTEM_DEFAULT) or
     (dwLCID = GetSystemDefaultLCID) or
     (dwLCID = GetUserDefaultLCID) then{cf 1.01.1}
  begin
    FLCID:= dwLcid;
    Result:= S_OK
  end else
  begin
    Result:= E_INVALIDARG
  end
end;

procedure TServerImpl.SinkConnect(const Sink: IUnknown;
  Connecting: Boolean);
begin
  if Connecting then
    FOPCShutdown:= Sink as IOPCShutdown
  else
    FOPCShutdown:= nil
end;

procedure TServerImpl.ForceDisconnect;
var
  g: TGroupImpl;
begin
  FConnectionPoint.FSink:= nil;
  FOpcShutdown:= nil;
  while FGroupList.Count > 0 do
  begin
    g:= FGroupList.Group(0);
    g.ForceDisconnect;
    RemoveGroup(Cardinal(g), true)
  end;
  CoDisconnectObject(Self, 0) {note that this will result in a call to self.Destroy}
end;

function TServerImpl.ObjQueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if GOpcItemServer.NoBrowsing and IsEqualGUID(IID, IOPCBrowseServerAddressSpace) then
    Result:= E_NOINTERFACE
  else
    Result:= inherited ObjQueryInterface(IID, Obj)
end;

{ TItemEnumerator }

constructor TItemEnumerator.Create(const aGroupItemList: IGroupItemList);
begin
 inherited Create;
 FIndex:= 0;
 GroupItemList:= aGroupItemList
end;

function TItemEnumerator.Next(celt: ULONG; out ppItemArray: POPCITEMATTRIBUTESARRAY;
                           pceltFetched: PULONG): HResult;
var
  i: Integer;
  ItemCount: Integer;
  icelt: Integer;
  ItemList: TGroupItemList;
begin
  {how many are available to fetch?}
  icelt:= celt;
  ItemList:= GroupItemList.List;
  ItemCount:= ItemList.Count - FIndex;
  if icelt > ItemCount then
  begin
    Result:= S_FALSE
  end else
  begin
    ItemCount:= icelt;
    Result:= S_OK
  end;
  if Assigned(pceltFetched) then
    pceltFetched^:= ItemCount;
  if ItemCount > 0 then
  begin
    ppItemArray:= CoTaskMemAlloc(SizeOf(OPCITEMATTRIBUTES)*ItemCount);
    i:= 0;
    while i < ItemCount do
    begin
      ItemList[FIndex].AssignTo(ppItemArray^[i]);
      Inc(FIndex);
      Inc(i)
    end
  end else
  begin
    ppItemArray:= nil
  end
end;

function TItemEnumerator.Skip(celt: Cardinal): HResult;
var
  List: TList;
begin
  List:= GroupItemList.List;
  if (FIndex + Integer(celt)) <= List.Count then
  begin
    Inc(FIndex, celt);
    Result:= S_OK
  end else
  begin
    FIndex:= List.Count;
    Result:= S_FALSE
  end
end;

function TItemEnumerator.Reset: HResult;
begin
  FIndex:=0;
  Result:= S_OK
end;

function TItemEnumerator.Clone(out ppEnumItemAttributes: IEnumOPCItemAttributes): HResult;
begin
  ppEnumItemAttributes:= TItemEnumerator.Create(GroupItemList);
  Result:= S_OK
end;

{ TGroupInfo }

function TGroupInfo.DA2Connected: Boolean;
begin
  Result:= Assigned(FDataCallback)
end;

function TGroupInfo.DA1Connected(Da1Format: TDa1Format): Boolean;
begin
  Result:= Assigned(FDa1Advise[Da1Format])
end;

{ TGroupImpl }

function TGroupImpl.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if GetInterface(IID, Obj) then
    Result:= S_OK
  else
    Result:= E_NOINTERFACE
end;

function TGroupImpl._AddRef: Integer;
begin
  Inc(ReferenceCount);
  Result:= ReferenceCount
end;

function TGroupImpl._Release: Integer;
begin
  Dec(ReferenceCount);
  Result:= ReferenceCount;
  if (ReferenceCount = 0) and (FDeleted) then
    Free 
end;

function TGroupImpl.AddItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  out ppAddResults: POPCITEMRESULTARRAY;
  out ppErrors: PResultList): HResult;
{$IFDEF GLD}
var
  i: Integer;
{$ENDIF}
begin
  Result:= AddOrValidate(True, False, dwCount, pItemArray, ppAddResults, ppErrors)
{$IFDEF GLD}
  ; SendDebug(Format('Added %d items', [dwCount]));
  if dwCount > 0 then
    for i:= 0 to dwCount - 1 do
      SendDebug(Format('  Item %d Server handle = %d', [i + 1, ppAddResults^[i].hServer]))
{$ENDIF}
end;

function TGroupImpl.AddOrValidate( DoAdd, BlobUpdate: Boolean;
  dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  var ppResults: POPCITEMRESULTARRAY;
  var ppErrors: PResultList): HResult;
var
  i: Integer;
  Items: POPCITEMDEFARRAY absolute pItemArray;
  ItemID: String;
  ItemResult: TServerItemRef;
  NewItem: TGroupItemImpl;
  IsError: Boolean;
begin
  ppResults:= nil;
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckCount(dwCount, pItemArray);
    ppResults:= ZeroAllocation(dwCount*sizeof(OPCITEMRESULT)); {cf 1.12.3}
    ppErrors:= CheckAllocation(dwCount*sizeof(HRESULT));
    IsError:= False;
    for i:= 0 to dwCount -1 do
    begin
      ppErrors^[i]:= S_OK;
      try
        if DoAdd then
        begin
          NewItem:= TGroupItemImpl.Create(Self, Items^[i], ppResults^[i]);
          FItemList.Add(NewItem);
        end else
        begin
          InitItemID(Items^[i], ItemID);
          ItemResult:= TServerItemRef.GetItem(ItemID);
          try
            ReadItemResult(ItemResult, ppResults^[i])
          finally
            ItemResult.ReleaseNonGroupReference
          end
        end;
      except
        on E: EOpcError do
          ppErrors^[i]:= E.ErrorCode
      end;
      if ppErrors^[i] <> S_OK then
        IsError := True;
    end;
    if not IsError then
      Result:= S_OK
    else
      Result:= S_FALSE
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
      FreeAndNull(ppResults);
      FreeAndNull(ppErrors)
    end
  end
end;

function TGroupImpl.AsyncIO2Read(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  dwTransactionID: DWORD; out pdwCancelID: DWORD;
  out ppErrors: PResultList): HResult;
{if you alter this to return an error array don't forget to free and null
if there is an exception &&&}
begin
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckConnected;
    CheckCount(dwCount, phServer);  {cf 1.13.8}
    if GOpcItemServer.AlwaysAllocateErrorArrays then
      ppErrors:= AllocErrorArray(dwCount); {cf. 1.12.1}
{$IFDEF GLD}
    SendDebug(Format('AsyncIO2Read dwCount = %d, Transid=%d', [dwCount, dwTransactionId]));
{$ENDIF}
    pdwCancelID:= DWORD(TAsyncReadTask.Create(Self, dwTransactionID, dwCount, phServer, ppErrors, Result));
{$IFDEF SLOWCALLBACKS}
    TestPostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
{$ELSE}
    PostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
{$ENDIF}
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end
end;

function TGroupImpl.AsyncIO2Write(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pItemValues: POleVariantArray; dwTransactionID: DWORD; out pdwCancelID: DWORD;
  out ppErrors: PResultList): HResult;
begin
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckConnected;
    CheckCount(dwCount, phServer);   {cf 1.13.9}
    if GOpcItemServer.AlwaysAllocateErrorArrays then
      ppErrors:= AllocErrorArray(dwCount); {cf. 1.12.1}
    pdwCancelID:= DWORD(TAsyncWriteTask.Create(Self, dwTransactionID, dwCount, phServer, pItemValues, ppErrors, Result));
    PostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end
end;

function TGroupImpl.DeleteTask(Task: TASyncTask): Boolean;
var
  i: Integer;
begin
  i:= FTaskList.IndexOf(Task);
  Result:= i <> -1;
  if Result then
  begin
    FTaskList.Delete(i);
    Task.Deleted:= true
  end
end;

function TGroupImpl.Cancel2(dwCancelID: DWORD): HResult;
begin
  try
{$IFDEF GLD}
    SendDebug(Format('AsyncTask.Cancel2 %p', [Pointer(dwCancelId)]));
{$ENDIF}
    CheckDeleted;
    if DeleteTask(TAsyncTask(dwCancelID)) then
      Result:= S_OK
    else
      Result:= E_FAIL
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

procedure TGroupImpl.CheckDeleted;
begin
  if FDeleted then
    raise EOpcError.Create(E_FAIL)
end;

procedure TGroupImpl.CheckConnected;
begin
  if not Assigned(FDataCallback) then
    raise EOpcError.Create(CONNECT_E_NOCONNECTION)
end;


function TGroupImpl.CloneGroup(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult;
var
  NewGroup: TGroupImpl;
  i: Integer;
  NewItem: TGroupItemImpl;
  List: TGroupList;
begin
  NewGroup:= nil;
  ppUnk:= nil;
  try
    CheckDeleted;
    NewGroup:= TGroupImpl.Create(FClientInfo,
                                szName,
                                False,
                                UpdateRate,
                                hClientGroup,
                                TimeBias,
                                FPercentDeadband,
                                FLCID);
    if not NewGroup.GetInterface(riid, ppUnk) then  {cf 1.13.7}
      raise EOpcError.Create(E_NOINTERFACE);
    for i:= 0 to FItemList.Count - 1 do
    begin
      NewItem:= TGroupItemImpl.CreateClone(NewGroup, FItemList[i]);
      NewGroup.FItemList.Add(NewItem)
    end;
    Result:= S_OK
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
      if Assigned(NewGroup) then
      begin  {cf 1.13.4}
        List:= TServerImpl(FClientInfo).FGroupList;
        i:= List.IndexOf(NewGroup.Name);
        if i <> -1 then
          List.Delete(i);
        NewGroup.Free
      end
    end
  end
end;

constructor TGroupImpl.Create( aServer: TClientInfo {IServerInternal};
                    const aName: String;
                    aActive: Boolean;
                    aUpdateRate: DWORD;
                    ahClientGroup: OPCHANDLE;
                    aTimeBias: Longint;
                    aPercentDeadband: Single;
                    aLCID: DWORD);
var
  FServer: TServerImpl absolute aServer;
begin
  inherited Create;
  FClientInfo:= aServer;
  FName:= Trim(aName); {cf 1.01.11}
  if FName = '' then
    FName:= FServer.FGroupList.GetUniqueGroupName;
{$IFDEF GLD}
   SendDebug(Format('GLD: Create %s RC = %d', [Name, ReferenceCount]));
{$ENDIF}
  SetPercentDeadband(aPercentDeadband);
  FDataChangeEnable:= true;
  FServer.FGroupList.AddGroup(FName, Self);
  FConnectionPoint:= TConnectionPoint.Create(Self, IOPCDataCallback, SinkConnect);
  FTaskList:= TList.Create;
  FItemList:= TGroupItemList.Create;
  SetActive(aActive);
  FhClientGroup:= ahClientGroup;
  TimeBias:= aTimeBias;
  FLCID:= aLCID;
  {start timer}
  FUpdateRate:= MaxInt;
  SetUpdateRate(aUpdateRate);
  {notify in main thread}
  SendMessage(OpcWindow, CM_NOTIFICATION, nfAddGroup, Integer(Self));
  Include(FGroupState, gsNotified) 
end;

procedure TGroupImpl.SetUpdateRate( Value: DWORD);
var
  MaxUpdate: DWORD;
begin
  if FUpdateRate <> Value then
  begin
    if gsTimerRunning in FGroupState then
    begin
      KillTimer(OpcWindow, Integer(Self));
      Exclude(FGroupState, gsTimerRunning)
    end;
    MaxUpdate:= GOpcItemServer.MaxUpdateRate;
    if Value < MaxUpdate then
      Value:= MaxUpdate;
    FUpdateRate:= Value;
    SetActive(FActive);
  end
end;

function TGroupImpl.CreateEnumerator(const riid: TIID;
  out ppUnk: IUnknown): HResult;
begin
  ppUnk:= nil;
  try
    CheckDeleted;
    if not IsEqualGUID(riid, IEnumOpcItemAttributes) then
      raise EOpcError.Create(E_INVALIDARG);
    ppUnk:= TItemEnumerator.Create(Self);
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

destructor TGroupImpl.Destroy;
var
  i: Integer;
  Msg: TMsg;
begin
{$IFDEF GLD}
   SendDebug(Format('GLD: Destroy %s RC = %d', [Name, ReferenceCount]));
{$ENDIF}
  if gsTimerRunning in FGroupState then
  begin
    KillTimer(OpcWindow, DWORD(Self));
    Exclude(FGroupState, gsTimerRunning);
    {remove all timer messages. These may not belong to this group,
    but the odd missed tick won't hurt. The chances of finding any
    are remote anyway}
    while PeekMessage(Msg, OpcWindow, WM_TIMER, WM_TIMER, PM_REMOVE) do ;

  end;
  if FGroupKeepAlive <> nil then
  begin
    FreeAndNil(FGroupKeepAlive);
    while PeekMessage(Msg, OpcWindow, WM_TIMER, WM_TIMER, PM_REMOVE) do ;
  end;
  {notify in main thread}
  if gsNotified in FGroupState then
  begin
    SendMessage(OpcWindow, CM_NOTIFICATION, nfRemoveGroup, Integer(Self));
    Exclude(FGroupState, gsNotified)
  end;
  FItemList.Free;
  FConnectionPoint.Free;
  if Assigned(FTaskList) then
  begin
    for i:= 0 to FTaskList.Count - 1 do
      TAsyncTask(FTaskList[i]).Deleted:= true;
    FTaskList.Free
  end;
  inherited Destroy
end;

procedure TGroupImpl.DoRefresh2( Source: OPCDATASOURCE; TransactionID: DWORD; ActiveList: TList);
var
  i: Integer;
  Count: Integer;
  ClientHandles: array of OPCHANDLE;
  Values: array of OleVariant;
  Qualities: array of Word;
  Timestamps: array of TFiletime;
  Errors: array of HRESULT;
  MasterError, MasterQuality: HRESULT;

begin
  if Assigned(FDataCallback) and FActive then
  begin
    Count:= ActiveList.Count;
    SetLength(ClientHandles, Count);
    SetLength(Values, Count);
    SetLength(Qualities, Count);
    SetLength(Timestamps, Count);
    SetLength(Errors, Count);
    MasterError:= S_OK;
    MasterQuality:= S_OK;
    for i:= 0 to Count - 1 do
    with TGroupItemImpl(ActiveList[i]) do
    begin
      Tick; // Clear cache changed
      ClientHandles[i]:= FhClient;
      try
        if Source = OPC_DS_DEVICE then
          GetItemValue(Values[i], Qualities[i], Timestamps[i])
        else
          GetCacheValue(Values[i], Qualities[i], Timestamps[i]);
        Errors[i]:= S_OK;
      except
        on E: EOpcError do
          Errors[i]:= E.ErrorCode;
      end;
      if (MasterError = S_OK) and (Errors[i] <> S_OK) then
        MasterError:= S_FALSE;
      if (MasterQuality = S_OK) and (Qualities[i] <> OPC_QUALITY_GOOD) then
        MasterQuality:= S_FALSE
    end;
    try   {try except cf 1.14.1}
      FDataCallback.OnDataChange(TransactionID,
       hClientGroup, MasterQuality, MasterError, Count, Pointer(ClientHandles),
       Pointer(Values), Pointer(Qualities), Pointer(Timestamps), Pointer(Errors));
      TServerImpl(FClientInfo).SetLastUpdate;
      if FGroupKeepAlive <> nil then
        FGroupKeepAlive.RestartTimer;
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Self, ccDa2DataChange)
    end
  end
end;

procedure TGroupImpl.DoRefreshMaxAge(MaxAge: DWORD; ActualTimestamp: TFileTime; TransactionID: DWORD; ActiveList: TList);
var
  i: Integer;
  Count: Integer;
  ClientHandles: array of OPCHANDLE;
  Values: array of OleVariant;
  Qualities: array of Word;
  Timestamps: array of TFiletime;
  Errors: array of HRESULT;
  MasterError, MasterQuality: HRESULT;
begin
  if Assigned(FDataCallback) and FActive then
  begin
    Count:= ActiveList.Count;
    SetLength(ClientHandles, Count);
    SetLength(Values, Count);
    SetLength(Qualities, Count);
    SetLength(Timestamps, Count);
    SetLength(Errors, Count);
    MasterError:= S_OK;
    MasterQuality:= S_OK;
    for i:= 0 to Count - 1 do
    with TGroupItemImpl(ActiveList[i]) do
    begin
      Tick; // Clear cache changed
      ClientHandles[i]:= FhClient;
      try
        GetMaxAgeValue(MaxAge, ActualTimestamp, Values[i], Qualities[i], Timestamps[i]);
        Errors[i]:= S_OK;
      except
        on E: EOpcError do
          Errors[i]:= E.ErrorCode;
      end;
      if (MasterError = S_OK) and (Errors[i] <> S_OK) then
        MasterError:= S_FALSE;
      if (MasterQuality = S_OK) and (Qualities[i] <> OPC_QUALITY_GOOD) then
        MasterQuality:= S_FALSE
    end;
    try   {try except cf 1.14.1}
      FDataCallback.OnDataChange(TransactionID,
       hClientGroup, MasterQuality, MasterError, Count, Pointer(ClientHandles),
       Pointer(Values), Pointer(Qualities), Pointer(Timestamps), Pointer(Errors));
      TServerImpl(FClientInfo).SetLastUpdate;
      if FGroupKeepAlive <> nil then
        FGroupKeepAlive.RestartTimer;
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Self, ccDa2DataChange)
    end
  end
end;

procedure TGroupImpl.DoRefresh1( Source: OPCDATASOURCE; TransactionID: DWORD;
          Format: TDa1Format; ActiveList: TList);
var
  Value: OleVariant;
  stgmed: TStgMedium;
  Fe: TFormatEtc;
  Count: Integer;
  Stream: TDataChangeStream;
  Timestamp: TFiletime;
  Status: HRESULT;
  i: Integer;
  GroupItem: TGroupItemImpl;
begin
  if Assigned(FDa1Advise[Format]) and FActive then
  begin
    with Fe do
    begin
      cfFormat:= GDa1Format[Format];
      ptd:= nil;
      dwAspect:= DVASPECT_CONTENT;
      lindex:= -1;
      tymed:= TYMED_HGLOBAL
    end;
    Count:= ActiveList.Count;
    if Format = da1DataTime then
      Stream:= TDataChangeStream1.Create(Count)
    else
      Stream:= TDataChangeStream2.Create(Count);
    try
      with Stream.GroupHeader^ do
      begin
        dwItemCount:= Count;
        hClientGroup:= Self.hClientGroup;
        dwTransactionID:= TransactionID
      end;
      Status:= S_OK;
      for i:= 0 to Count - 1 do
      begin
        GroupItem:= TGroupItemImpl(ActiveList[i]);
        with Stream.ItemHeader2(i)^ do
        begin
          hClient:= GroupItem.FhClient;
          wReserved:= 0;
          dwValueOffset:= Stream.Position; // + sizeof(OPCITEMHEADER2);
          Value:= Null;
          try
            if Source = OPC_DS_DEVICE then
              GroupItem.GetItemValue(Value, wQuality, Timestamp)
            else
              GroupItem.GetCacheValue(Value, wQuality, Timestamp);
            Stream.SetTimestamp(i, Timestamp)
          except
            on EOpcError do
              wQuality:= OPC_QUALITY_BAD
          end;
          if (Status = S_OK) and (wQuality <> OPC_QUALITY_GOOD) then
            Status:= S_FALSE
        end;
        {note that any writes to the stream may alter the
        base pointer, therefore we MUST NOT write to the
        stream inside with Stream.ItemHeader1(i)^ do}
        Stream.WriteVariant(Value) {cf 1.13.12}
      end;
      with Stream.GroupHeader^ do
      begin
        dwSize:= Stream.Position;
        hrStatus:= Status
      end;
      stgmed.tymed:= TYMED_HGLOBAL;
      stgmed.hGlobal:= GlobalHandle(Stream.Memory);
      stgmed.UnkForRelease:= nil;
      if Assigned(FDa1Advise[Stream.Format]) then
      begin
        try   {try except cf 1.14.6}
          FDa1Advise[Stream.Format].OnDataChange(
            fe, stgmed)
        except
          on E: EOleSysError do
            GOpcItemServer.GroupCallbackError(E, Self, ccDa1DataChange)
        end
      end
    finally
      Stream.Free
    end
  end
end;

function TGroupImpl.GetEnable(out pbEnable: BOOL): HResult;
var
  IntEnable: Cardinal absolute pbEnable;
begin
  try
    CheckDeleted;
    CheckConnected;
    if FDataChangeEnable then  {cf 1.14.9}
      IntEnable:= 1
    else
      IntEnable:= 0;
{    pbEnable:= FDataChangeEnable; }
{$IFDEF GLD}
    if FDataChangeEnable then
      SendDebug('Get enable TRUE')
    else
      SendDebug('Get enable FALSE');
{$ENDIF}
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.AsyncIO3ReadMaxAge(dwCount: DWORD;
  phServer: POPCHANDLEARRAY; pdwMaxAge: PDWORDARRAY;
  dwTransactionID: DWORD; out pdwCancelID: DWORD;
  out ppErrors: PResultList): HResult;
{if you alter this to return an error array don't forget to free and null
if there is an exception &&&}
begin
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckConnected;
    CheckCount(dwCount, phServer);  {cf 1.13.8}
    if GOpcItemServer.AlwaysAllocateErrorArrays then
      ppErrors:= AllocErrorArray(dwCount); {cf. 1.12.1}
{$IFDEF GLD}
    SendDebug(Format('AsyncIO3ReadMaxAge dwCount = %d, Transid=%d', [dwCount, dwTransactionId]));
{$ENDIF}
    pdwCancelID:= DWORD(TAsyncReadMaxAgeTask.Create(Self, dwTransactionID, dwCount, phServer, pdwMaxAge, ppErrors, Result));
{$IFDEF SLOWCALLBACKS}
    TestPostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
{$ELSE}
    PostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
{$ENDIF}
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end
end;

function TGroupImpl.AsyncIO3WriteVQT(dwCount: DWORD;
  phServer: POPCHANDLEARRAY; pItemVQT: POPCITEMVQTARRAY;
  dwTransactionID: DWORD; out pdwCancelID: DWORD;
  out ppErrors: PResultList): HResult;
begin
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckConnected;
    CheckCount(dwCount, phServer);   {cf 1.13.9}
    if GOpcItemServer.AlwaysAllocateErrorArrays then
      ppErrors:= AllocErrorArray(dwCount); {cf. 1.12.1}
    pdwCancelID:= DWORD(TAsyncWriteVQTTask.Create(Self, dwTransactionID, dwCount, phServer, pItemVQT, ppErrors, Result));
    PostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end
end;

function TGroupImpl.AsyncIO3RefreshMaxAge(dwMaxAge, dwTransactionID: DWORD;
  out pdwCancelID: DWORD): HResult;
begin
  try
    CheckDeleted;
    CheckConnected;
    if not FActive then
      raise EOpcError.Create(E_FAIL);  {Section 4.3.2 of DA2 spec}
    pdwCancelID:= DWORD(TRefreshMaxAgeTask.Create(Self, dwTransactionID, dwMaxAge));
    PostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.SetItemDeadband(dwCount: DWORD;
  phServer: POPCHANDLEARRAY; pPercentDeadband: PSingleArray;
  out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, pPercentDeadband, ItemSetDeadband, ppErrors)
end;

function TGroupImpl.GetItemDeadband(dwCount: DWORD;
  phServer: POPCHANDLEARRAY; out ppPercentDeadband: PSingleArray;
  out ppErrors: PResultList): HResult;
begin
  ppPercentDeadband:= ZeroAllocation(dwCount*SizeOf(Single));
  Result:= IterateItems(dwCount, phServer, ppPercentDeadband, ItemGetDeadband, ppErrors)
end;

function TGroupImpl.ClearItemDeadband(dwCount: DWORD;
  phServer: POPCHANDLEARRAY; out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, nil, ItemClearDeadband, ppErrors)
end;

function TGroupImpl.GetGroupItemList: TGroupItemList;
begin
  Result:= FItemList
end;

function TGroupImpl.GetState
(out pUpdateRate: DWORD;
  out pActive: BOOL; out ppName: POleStr; out pTimeBias: Integer;
  out pPercentDeadband: Single; out pLCID: TLCID; out phClientGroup,
  phServerGroup: OPCHANDLE): HResult;
begin
  ppName:= nil;
  try
    CheckDeleted;
    pUpdateRate:= UpdateRate;
    Integer(pActive):= ShortInt(FActive); {cf 1.01.2}
    ppName:= StringToLPOLESTR(FName);
    pTimeBias:= TimeBias;
    pPercentDeadband:= FPercentDeadband;
    pLCID:= FLCID;
    phClientGroup:= hClientGroup;
    phServerGroup:= OPCHANDLE(Self);
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.InScope(dwScope: OPCENUMSCOPE): Boolean;
begin
  case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,
    OPC_ENUM_PRIVATE: Result:= not FIsPublicGroup;
    OPC_ENUM_PUBLIC_CONNECTIONS,
    OPC_ENUM_PUBLIC: Result:= FIsPublicGroup;
    OPC_ENUM_ALL_CONNECTIONS,
    OPC_ENUM_ALL: Result:= true;
  else
    Result:= false
  end
end;

procedure TGroupImpl.ItemRemove(Item: TGroupItemImpl; Index: Integer;
  Data: Pointer);
var
  i: Integer;
begin
  i:= FItemList.IndexOf(Item);
  if i = -1 then
    raise EOpcError.Create(OPC_E_INVALIDHANDLE); {already checked &&&}
  FItemList.Delete(i)
end;

procedure TGroupImpl.ItemSetActiveState(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
var
  bActive: PBOOL absolute Data;
begin
  Item.SetActive(bActive^)
end;

procedure TGroupImpl.ItemSetClientHandle(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
var
  ClientHandles: POPCHANDLEARRAY absolute Data;
begin
  Item.FhClient:= ClientHandles^[Index]
end;

procedure TGroupImpl.ItemSetDatatype(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
var
  TypeList: PVarTypeList absolute Data;
  NewType: TVarType;
begin
  NewType:= TypeList^[Index];
  if not IsSupportedVarType(NewType) then
    raise EOpcError.Create(OPC_E_BADTYPE);

  Item.FRequestedDataType:= NewType;
  Item.InvalidateCache;
end;

type
  PSyncIOParam = ^TSyncIOParam;
  TSyncIOParam = record
    GroupActive: Boolean;
    Source: OPCDATASOURCE;
    ppItemValues: POPCITEMSTATEARRAY;
  end;

  PSyncIO2Param = ^TSyncIO2Param;
  TSyncIO2Param = record
    GroupActive: Boolean;
    ActualTimestamp: TFileTime;
    pdwMaxAge: PDWORDARRAY;
    ppvValues: POleVariantArray;
    ppwQualities: PWordArray;
    ppftTimeStamps: PFileTimeArray;
  end;

  PAsyncReadParam = ^TAsyncReadParam;
  TAsyncReadParam = record
    hClientItem: array of OPCHANDLE;
    Qualities: array of Word;
    Timestamps: array of TFiletime;
    Values: array of OleVariant;
  end;

  PAsyncReadMaxAgeParam = ^TAsyncReadMaxAgeParam;
  TAsyncReadMaxAgeParam = record
    hClientItem: array of OPCHANDLE;
    Qualities: array of Word;
    Timestamps: array of TFiletime;
    Values: array of OleVariant;
  end;

procedure TGroupImpl.ItemSyncIORead(Item: TGroupItemImpl; Index: Integer;
  Data: Pointer);
var
  IOP: PSyncIOParam absolute Data;
begin
  if (Item.ServerItem.AccessRights and OPC_READABLE) = 0 then
    raise EOpcError.Create(OPC_E_BADRIGHTS);
  with IOP^.ppItemValues^[Index] do
  begin
    hClient:= Item.hClient;     {cf 1.13.1}
    if IOP^.Source = OPC_DS_CACHE then
    begin
      Item.GetCacheValue(vDataValue, wQuality, ftTimestamp);
      if not (Item.FActive and IOP^.GroupActive) then
        wQuality:= OPC_QUALITY_OUT_OF_SERVICE
    end else
    begin
      Item.GetItemValue(vDataValue, wQuality, ftTimestamp)
    end;
    wReserved:= 0
  end
end;

procedure TGroupImpl.ItemSyncIOWrite(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
var
  pItemValues: POleVariantArray absolute Data;
begin
  Item.ServerItem.SetItemValue(pItemValues^[Index]);
end;

procedure TGroupImpl.ItemSyncIO2Read(Item: TGroupItemImpl; Index: Integer;
  Data: Pointer);
var
  IOP: PSyncIO2Param absolute Data;
begin
  if (Item.ServerItem.AccessRights and OPC_READABLE) = 0 then
    raise EOpcError.Create(OPC_E_BADRIGHTS);
  Item.GetMaxAgeValue(IOP^.pdwMaxAge[Index], IOP^.ActualTimestamp,
    IOP^.ppvValues[Index], IOP^.ppwQualities[Index], IOP^.ppftTimeStamps[Index]);
end;

procedure TGroupImpl.ItemSyncIO2Write(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
var
  pItemValues: POPCITEMVQTARRAY absolute Data;
begin
  Item.ServerItem.SetItemVQT(pItemValues^[Index]);
end;

procedure TGroupImpl.ItemSetDeadband(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
const
  Eps = 1E-20;
var
  pPercentDeadband: PSingleArray absolute Data;
  Value: Single;
begin
  Value:= pPercentDeadband^[Index];
  if (Value < -Eps) or (Value > 100+Eps) then
    raise EOpcError.Create(E_INVALIDARG);

  Item.PercentDeadband:= Value;
  Item.PercentDeadbandSet:= True;
end;

procedure TGroupImpl.ItemGetDeadband(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
var
  pPercentDeadband: PSingleArray absolute Data;
begin
  if not Item.PercentDeadbandSet then
    raise EOpcError.Create(OPC_E_DEADBANDNOTSET);
  pPercentDeadband^[Index]:= Item.PercentDeadband;
end;

procedure TGroupImpl.ItemClearDeadband(Item: TGroupItemImpl;
  Index: Integer; Data: Pointer);
begin
  if not Item.PercentDeadbandSet then
    raise EOpcError.Create(OPC_E_DEADBANDNOTSET);
  Item.PercentDeadband := 0;
  Item.PercentDeadbandSet := False;
end;

function TGroupImpl.IterateItems(dwCount: Integer;
  phServer: POPCHANDLEARRAY; Data: Pointer; Action: TItemAction;
  var ppErrors: PResultList): HRESULT;
var
  i: Integer;
  IsError: Boolean;
begin
  ppErrors:= nil;
  try
    CheckCount(dwCount, phServer);
    ppErrors:= CheckAllocation(dwCount*SizeOf(HRESULT));
    IsError:= False;
    for i:= 0 to dwCount - 1 do
    begin
      try
        ValidateServerHandle(phServer^[i]);
        Action(TGroupItemImpl(phServer^[i]), i, Data);
        ppErrors^[i]:= S_OK
      except
        on E: EOpcError do
        begin
          IsError:= True;
          ppErrors^[i]:= E.ErrorCode
        end
      end;
    end;
    if not IsError then
      Result:= S_OK
    else
      Result:= S_FALSE
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
      FreeAndNull(ppErrors)
    end
  end
end;

function TGroupImpl.Refresh2(dwSource: OPCDATASOURCE;
  dwTransactionID: DWORD; out pdwCancelID: DWORD): HResult;
begin
  pdwCancelID:= 0;
  try
    CheckDeleted;
    CheckConnected;
    TestParamRange(dwSource, OPC_DS_CACHE, OPC_DS_DEVICE);
    if not FActive then
      raise EOpcError.Create(E_FAIL);  {Section 4.3.2 of DA2 spec}
    pdwCancelID:= DWORD(TRefreshTask.Create(Self, dwTransactionID, dwSource));
    PostMessage(OpcWindow, CM_ASYNCTASK, pdwCancelID, 0);
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.RemoveItems(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, nil, ItemRemove, ppErrors)
{$IFDEF GLD}
  ; SendDebug(Format('Removed %d items', [dwCount]))
{$ENDIF}
end;

function TGroupImpl.SetActiveState(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  bActive: BOOL; out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, @bActive, ItemSetActiveState, ppErrors)
end;

function TGroupImpl.SetClientHandles(dwCount: DWORD; phServer,
  phClient: POPCHANDLEARRAY; out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, phClient, ItemSetClientHandle, ppErrors)
end;

function TGroupImpl.SetDatatypes(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pRequestedDatatypes: PVarTypeList; out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, pRequestedDatatypes, ItemSetDatatype, ppErrors)
end;

function TGroupImpl.SetEnable(bEnable: BOOL): HResult;
var
  CancelID: DWORD;
begin
  try
    CheckDeleted;
    CheckConnected;
    FDataChangeEnable:= bEnable;
    if bEnable then
      Refresh2(OPC_DS_CACHE, 0, CancelID);
{$IFDEF GLD}
    if FDataChangeEnable then
      SendDebug('Set enable TRUE')
    else
      SendDebug('Set enable FALSE');
{$ENDIF}
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.SetName(szName: POleStr): HResult;
var
  Str: String;

procedure InvalidArg;
begin
  raise EOpcError.Create(E_INVALIDARG)
end;

begin
  try
    if FDeleted then
    begin
      Result:= E_FAIL
    end else
    begin
      if not Assigned(szName) then
        InvalidArg;
      Str:= Trim(szName);
      if Str = '' then
        InvalidArg;
      if TServerImpl(FClientInfo).FGroupList.FindGroup(Str) then
        raise EOpcError.Create(OPC_E_DUPLICATENAME);
      FName:= Str;
      TServerImpl(FClientInfo).FGroupList.UpdateNames;
      Result:= S_OK
    end
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.SetState(pRequestedUpdateRate: PDWORD;
  out pRevisedUpdateRate: DWORD; pActive: PBOOL; pTimeBias: PLongint;
  pPercentDeadband: PSingle; pLCID: PLCID;
  phClientGroup: POPCHANDLE): HResult;
begin
  try
    CheckDeleted;
    if Assigned(pRequestedUpdateRate) then
    begin
      SetUpdateRate(pRequestedUpdateRate^);
      pRevisedUpdateRate:= UpdateRate
    end;
    if Assigned(pActive) then
      SetActive(pActive^);
    if Assigned(pTimeBias) then
      TimeBias:= pTimeBias^;
    if Assigned(pPercentDeadband) then
      SetPercentDeadband(pPercentDeadband^);  {cf 1.01.3}
    if Assigned(pLCID) then
      FLCID:= pLCID^;
    if Assigned(phClientGroup) then
      FhClientGroup:= phClientGroup^;
    if Assigned(PRequestedUpdateRate) and  {cf 1.01.10}
       (pRevisedUpdateRate <> pRequestedUpdateRate^) then
      Result:= OPC_S_UNSUPPORTEDRATE
    else
      Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

procedure TGroupImpl.SinkConnect(const Sink: Iunknown; Connecting: Boolean);
var
  CancelID: DWORD;
begin
  if Connecting then
  begin
    FDataCallback:= Sink as IOPCDataCallback;
    if FDataChangeEnable and
       FActive then
      Refresh2(OPC_DS_DEVICE, 0, CancelID);
  end else
  begin
    FDataCallback:= nil
  end
end;

function TGroupImpl.SyncIORead(dwSource: OPCDATASOURCE; dwCount: DWORD;
  phServer: POPCHANDLEARRAY; out ppItemValues: POPCITEMSTATEARRAY;
  out ppErrors: PResultList): HResult;
{
  PSyncIOParam = ^TSyncIOParam;
  TSyncIOParam = record
    GroupActive: Boolean;
    Source: OPCDATASOURCE;
    ppItemValues: POPCITEMSTATEARRAY;
  end;
}
var
  IOP: TSyncIOParam;
begin
  ppItemValues:= nil;
  try
    CheckCount(dwCount, phServer);
    TestParamRange(dwSource, OPC_DS_CACHE, OPC_DS_DEVICE);
    ppItemValues:= ZeroAllocation(SizeOf(OPCITEMSTATE)*dwCount); {cf 1.12.3}
    IOP.ppItemValues:= ppItemValues;
    IOP.Source:= dwSource;
    IOP.GroupActive:= FActive;
    Result:= IterateItems(dwCount, phServer, @IOP, ItemSyncIORead, ppErrors);
    if Succeeded(Result) then
      TServerImpl(FClientInfo).SetLastUpdate
    else
      FreeAndNull(ppItemValues)
        {in theory, I should raise exception here, but in practice
        it would be absurd}
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
      FreeAndNull(ppItemValues)
    end
  end
end;

function TGroupImpl.SyncIOWrite(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pItemValues: POleVariantArray; out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, pItemValues, ItemSyncIOWrite, ppErrors)
end;

function TGroupImpl.SyncIO2ReadMaxAge(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pdwMaxAge: PDWORDARRAY; out ppvValues: POleVariantArray;
  out ppwQualities: PWordArray; out ppftTimeStamps: PFileTimeArray;
  out ppErrors: PResultList): HResult;
var
  IOP: TSyncIO2Param;
begin
  try
    CheckCount(dwCount, phServer);
    ppvValues := ZeroAllocation(dwCount * SizeOf(OleVariant));
    ppwQualities := ZeroAllocation(dwCount * SizeOf(Word));
    ppftTimeStamps := ZeroAllocation(dwCount * SizeOf(TFileTime));

    IOP.ppvValues := ppvValues;
    IOP.ppwQualities := ppwQualities;
    IOP.ppftTimeStamps := ppftTimeStamps;
    IOP.GroupActive:= FActive;
    GOpcItemServer.GetTimestamp(IOP.ActualTimestamp);
    IOP.pdwMaxAge := pdwMaxAge;

    Result:= IterateItems(dwCount, phServer, @IOP, ItemSyncIO2Read, ppErrors);
    if Succeeded(Result) then
      TServerImpl(FClientInfo).SetLastUpdate
    else begin
      FreeAndNull(ppvValues);
      FreeAndNull(ppwQualities);
      FreeAndNull(ppftTimeStamps);
    end;
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
      FreeAndNull(ppvValues);
      FreeAndNull(ppwQualities);
      FreeAndNull(ppftTimeStamps);
    end
  end
end;

function TGroupImpl.SyncIO2WriteVQT(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pItemVQT: POPCITEMVQTARRAY; out ppErrors: PResultList): HResult;
begin
  Result:= IterateItems(dwCount, phServer, pItemVQT, ItemSyncIO2Write, ppErrors)
end;

function TGroupImpl.ValidateItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  bBlobUpdate: BOOL; out ppValidationResults: POPCITEMRESULTARRAY;
  out ppErrors: PResultList): HResult;
begin
  Result:= AddOrValidate(False, bBlobUpdate, dwCount, pItemArray, ppValidationResults, ppErrors)
end;

function TGroupImpl.EnumConnectionPoints(
  out enumconn: IEnumConnectionPoints): HResult;
begin
  enumconn:= TEnumerateCP.Create(FConnectionPoint, 0); {cf 1.01.5}
  Result:= S_OK
end;

function TGroupImpl.FindConnectionPoint(const iid: TIID;
  out cp: IConnectionPoint): HResult;
begin
  if IsEqualGUID(FConnectionPoint.FIID, iid) then
  begin
    cp:= FConnectionPoint;
    Result:= S_OK
  end else
  begin
    Result:= CONNECT_E_NOCONNECTION
  end
end;

procedure TGroupImpl.Tick;
{refer to Section 4.3.3 of spec}
var
  List: TList;
  i: Integer;
  gi: TGroupItemImpl;
  j: TDa1Format;
begin
  if FActive then   {otherwise nothing}
  begin
    List:= TList.Create;
    try
      for i:= 0 to FItemList.Count - 1 do
      begin
        gi:= FItemList[i];
        if gi.Tick then
          List.Add(gi)
      end;
      if List.Count > 0 then
      begin
        if Assigned(FDataCallback) and
           FDataChangeEnable then
          DoRefresh2(OPC_DS_CACHE, 0, List);
        for j:= da1Data to da1DataTime do
        if Assigned(FDa1Advise[j]) then
          DoRefresh1(OPC_DS_CACHE, 0, j, List)
      end
    finally
      List.Free
    end
  end
end;

procedure TGroupImpl.SetActive(Value: Boolean);
var
  i: Integer;
  CancelID: DWORD;
begin
  if FActive <> Value then
  begin
    FActive:= Value;
    if Value then
    begin
      if Assigned(FDataCallback) and
         FDataChangeEnable then
        Refresh2(OPC_DS_CACHE, 0, CancelID) //?? DEVICE or CACHE?
      else
        for i:= 0 to FItemList.Count - 1 do
          FItemList[i].InvalidateCache;
    end;
  end;

  if Value and not (gsTimerRunning in FGroupState) then
  begin
    if SetTimer(OpcWindow, DWORD(Self), UpdateRate, nil) <> 0 then
      Include(FGroupState, gsTimerRunning)
    else
      raise EOpcError.Create(E_FAIL);
  end;

  if not Value and (gsTimerRunning in FGroupState) then
  begin
    KillTimer(OpcWindow, Integer(Self));
    Exclude(FGroupState, gsTimerRunning)
  end;
end;


function TGroupImpl.AsyncIORead(dwConnection: DWORD;
  dwSource: OPCDATASOURCE; dwCount: DWORD; phServer: POPCHANDLEARRAY;
  out pTransactionID: DWORD; out ppErrors: PResultList): HResult;
begin
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckCount(dwCount, phServer);
    CheckDa1Connected(dwConnection);
    TestParamRange(dwSource, OPC_DS_CACHE, OPC_DS_DEVICE);
    if AsyncBadHandle(dwCount, phServer, ppErrors) then   {cf 1.13.11}
    begin
      pTransactionID:= 0;
      Result:= S_FALSE
    end else
    begin
      if GOpcItemServer.AlwaysAllocateErrorArrays then
        ppErrors:= AllocErrorArray(dwCount); {cf. 1.12.1}
      pTransactionID:=
        DWORD(TAsyncReadTask1.Create(Self, dwCount, phServer, dwSource,
          TDa1Format(dwConnection), ppErrors, Result));
      PostMessage(OpcWindow, CM_ASYNCTASK, pTransactionID, 0);
    end
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end
end;

function TGroupImpl.AsyncIORefresh(dwConnection: DWORD;
  dwSource: OPCDATASOURCE; out pTransactionID: DWORD): HResult;
begin
  pTransactionID:= 0;
  try
    CheckDeleted;
    CheckDa1Connected(dwConnection);
    TestParamRange(dwSource, OPC_DS_CACHE, OPC_DS_DEVICE);
    if not FActive then
      raise EOpcError.Create(E_FAIL);  {Section 4.3.2 of DA2 spec}
    pTransactionID:= DWORD(TRefreshTask1.Create(Self, dwSource, TDa1Format(dwConnection)));
    PostMessage(OpcWindow, CM_ASYNCTASK, pTransactionID, 0);
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.AsyncIOWrite(dwConnection, dwCount: DWORD;
  phServer: POPCHANDLEARRAY; pItemValues: POleVariantArray;
  out pTransactionID: DWORD; out ppErrors: PResultList): HResult;
begin
  ppErrors:= nil;
  try
    CheckDeleted;
    CheckCount(dwCount, phServer);
    CheckDa1Connected(dwConnection);
    if AsyncBadHandle(dwCount, phServer, ppErrors) then   {cf 1.13.11}
    begin
      pTransactionID:= 0;
      Result:= S_FALSE
    end else
    begin
      if GOpcItemServer.AlwaysAllocateErrorArrays then
        ppErrors:= AllocErrorArray(dwCount); {cf. 1.12.1}
      pTransactionID:= DWORD(TAsyncWriteTask1.Create(Self, dwCount, phServer, pItemValues, ppErrors, Result));
      PostMessage(OpcWindow, CM_ASYNCTASK, pTransactionID, 0);
    end
  except
    on E: EOpcError do
      Result:= E.ErrorCode;
  end
end;

function TGroupImpl.DAdvise(const formatetc: TFormatEtc; advf: Integer;
  const advSink: IAdviseSink; out dwConnection: Integer): HResult;
var
  i: TDa1Format;
  FoundFormat: Boolean;
  Format: TDa1Format;
  TransId: DWORD;
begin
  try
    dwConnection:= -1;
    Format:= Low(TDa1Format);
    with formatetc do
    begin
      if lindex <> -1 then
        raise EOpcError.Create(DV_E_LINDEX);
      if (dwAspect <> DVASPECT_CONTENT) or
         (ptd <> nil) or
         (tymed <> TYMED_HGLOBAL) then
        raise EOpcError.Create(DV_E_FORMATETC);
      FoundFormat:= false;
      for i:= Low(TDa1Format) to High(TDa1Format) do
      if cfFormat = GDa1Format[i] then
      begin
        Format:= i;
        FoundFormat:= true;
        break
      end;
      if not FoundFormat then
        raise EOpcError.Create(DV_E_FORMATETC);
      if Assigned(FDa1Advise[Format]) then
        raise EOpcError.Create(CONNECT_E_ADVISELIMIT);
      FDa1Advise[Format]:= advSink;
      dwConnection:= Integer(Format);
      AsyncIORefresh(dwConnection, OPC_DS_CACHE, Transid);
    end;
    Result:= S_OK
  except
    on E: EOpcError do
      Result:= E.ErrorCode
  end
end;

function TGroupImpl.DUnadvise(dwConnection: Integer): HResult;
begin
  if (dwConnection < Ord(Low(TDa1Format))) or
     (dwConnection > Ord(High(TDa1Format))) or
     not Assigned(FDa1Advise[TDa1Format(dwConnection)]) then
  begin
    Result:= OLE_E_NOCONNECTION
  end else
  begin
    FDa1Advise[TDa1Format(dwConnection)]:= nil;
    Result:= S_OK
  end
end;

function TGroupImpl.EnumDAdvise(out enumAdvise: IEnumStatData): HResult;
begin
  Result:= E_NOTIMPL
end;

function TGroupImpl.EnumFormatEtc(dwDirection: Integer;
  out enumFormatEtc: IEnumFormatEtc): HResult;
begin
  Result:= E_NOTIMPL
end;

function TGroupImpl.GetCanonicalFormatEtc(const formatetc: TFormatEtc;
  out formatetcOut: TFormatEtc): HResult;
begin
  Result:= E_NOTIMPL
end;

function TGroupImpl.GetData(const formatetcIn: TFormatEtc;
  out medium: TStgMedium): HResult;
begin
  Result:= E_NOTIMPL
end;

function TGroupImpl.GetDataHere(const formatetc: TFormatEtc;
  out medium: TStgMedium): HResult;
begin
  Result:= E_NOTIMPL
end;

function TGroupImpl.QueryGetData(const formatetc: TFormatEtc): HResult;
begin
  Result:= E_NOTIMPL
end;

function TGroupImpl.SetData(const formatetc: TFormatEtc;
  var medium: TStgMedium; fRelease: BOOL): HResult;
begin
  Result:= E_NOTIMPL
end;

procedure TGroupImpl.CheckDa1Connected(dwConnection: Integer);
begin
  if (dwConnection < Ord(Low(TDa1Format))) or
     (dwConnection > Ord(High(TDa1Format))) or
     not Assigned(FDa1Advise[TDa1Format(dwConnection)]) then
    raise EOpcError.Create(CONNECT_E_NOCONNECTION)
end;

function TGroupImpl.Item(i: Integer): TGroupItemInfo;
begin
  Result:= FItemList[i]
end;

function TGroupImpl.ItemCount: Integer;
begin
  Result:= FItemList.Count
end;

procedure TGroupImpl.ValidateServerHandle(hServer: OPCHANDLE);
var
  Obj: TObject absolute hServer;
  i: Integer;
begin  {cf 1.14.10}
  if GOpcItemServer.StrictHandleValidation then  {this is slow because it requries a linear search of
                           the item list}
  begin
    i:= FItemList.IndexOf(Obj);
    if i = -1 then
      raise EOpcError.Create(OPC_E_INVALIDHANDLE)
  end else
  begin
    try
      if Obj.ClassType <> TGroupItemImpl then
        raise EOpcError.Create(OPC_E_INVALIDHANDLE)
    except
      on EAccessViolation do
        raise EOpcError.Create(OPC_E_INVALIDHANDLE)
    end
  end
end;

procedure TGroupImpl.SetPercentDeadband(Value: Single);
const
  Eps = 1E-20;
begin
  if (Value < -Eps) or (Value > 100+Eps) then
    raise EOpcError.Create(E_INVALIDARG);
  FPercentDeadband:= Value
end;

function TGroupImpl.AllocErrorArray(dwCount: DWORD): PResultList; {cf. 1.12.1}
begin
  Result:= ZeroAllocation(dwCount*SizeOf(HRESULT))  {cf 1.12.3}
end;

function TGroupImpl.ASyncBadHandle(dwCount: DWORD;
  phServer: POPCHANDLEARRAY; var ppErrors: PResultList): Boolean;
var
  i: Integer;
  Temp: array of HResult;
begin
  ppErrors:= nil;
  Result:= false;
  SetLength(Temp, dwCount);
  for i:= 0 to dwCount - 1 do
  begin
    try
      ValidateServerHandle(phServer^[i]);
      Temp[i]:= 0;
    except
      on E: EOpcError do
      begin
        Temp[i]:= E.ErrorCode;
        Result:= true
      end
    end;
    if Result then
    begin
      ppErrors:= CheckAllocation(dwCount*SizeOf(HRESULT));
      Move(Temp[0], ppErrors^, dwCount*SizeOf(HRESULT))
    end
  end
end;

procedure TGroupImpl.ForceDisconnect;
var
  i: TDa1Format;
begin
  FConnectionPoint.FSink:= nil;
  FDataCallback:= nil;
  for i:= Low(FDa1Advise) to High(FDa1Advise) do
    FDa1Advise[i]:= nil;
  CoDisconnectObject(Self, 0)
end;

function TGroupImpl.GetKeepAlive(out pdwKeepAliveTime: DWORD): HResult;
begin
  if FGroupKeepAlive = nil then
    pdwKeepAliveTime := 0
  else
    pdwKeepAliveTime := FGroupKeepAlive.KeepAlive;
  Result := S_OK;
end;

function TGroupImpl.SetKeepAlive(dwKeepAliveTime: DWORD;
  out pdwRevisedKeepAliveTime: DWORD): HResult;
begin
  if dwKeepAliveTime = 1 then
  begin
    Result := OPC_S_UNSUPPORTEDRATE;
    Exit;
  end;

  pdwRevisedKeepAliveTime := dwKeepAliveTime;

  if FGroupKeepAlive <> nil then
    FreeAndNil(FGroupKeepAlive);

  if dwKeepAliveTime <> 0 then
  begin
    FGroupKeepAlive := TGroupKeepAlive.Create(Self, dwKeepAliveTime);
    pdwRevisedKeepAliveTime := FGroupKeepAlive.KeepAlive;
  end;

  Result := S_OK;
end;

{ TGroupKeepAlive }

constructor TGroupKeepAlive.Create(AGroup: TGroupImpl; AKeepAlive : DWORD);
begin
  Group := AGroup;
  KeepAlive := AKeepAlive;
  if SetTimer(OpcWindow, DWORD(Self), KeepAlive, nil) = 0 then
    raise EOpcError.Create(E_FAIL);
end;

destructor TGroupKeepAlive.Destroy;
begin
  KillTimer(OpcWindow, DWORD(Self));
  inherited;
end;

procedure TGroupKeepAlive.RestartTimer;
begin
  KillTimer(OpcWindow, DWORD(Self));
  if SetTimer(OpcWindow, DWORD(Self), KeepAlive, nil) = 0 then
    raise EOpcError.Create(E_FAIL);
end;

procedure TGroupKeepAlive.Tick;
begin
  try
    Group.FDataCallback.OnDataChange(0, Group.hClientGroup,
      OPC_QUALITY_GOOD, S_OK, 0, nil, nil, nil, nil, nil);
  finally
  end;
end;

{ TServerItemRef }

procedure TServerItemRef.AssignTo(var Result: OPCITEMATTRIBUTES);
begin
  Result.szAccessPath:= StringToLPOLEStr('');
  Result.szItemID:= StringToLPOLESTR(ItemID);
  Result.dwAccessRights:= AccessRights;
  Result.vtCanonicalDataType:= CanonicalDataType;
  if Assigned(EUInfo) then
  begin
    Result.dwEUType:= Integer(EUInfo.EUType);
    Result.vEUInfo:= EUInfo.EUInfo
  end
end;

function NumericType(Vt: TVarType): Boolean;
begin
  case Vt of
    VT_UI1, VT_I2, VT_I4, VT_R4, VT_R8, VT_CY, VT_I1, VT_UI2, VT_UI4, VT_INT,
    VT_UINT: Result:= true;
  else
    Result:= false
  end
end;

constructor TServerItemRef.Create(const aItemID: String);
var
  Ar: TAccessRights;
  AccessPath: String;
begin
  inherited Create;
  ItemHandle:= INVALID_HANDLE_VALUE;
  Ar:= AllAccess;
  ItemID:= aItemID;
  AccessPath:= '';
  ItemHandle:= GOpcItemServer.GetExtendedItemInfo(ItemID, AccessPath,
    Ar, EUInfo, ItemProperties);
  AccessRights:= OpcAccessRights(Ar);
  CacheQuality:= OPC_QUALITY_GOOD;
  CacheValue:= GOpcItemServer.GetItemVQT(ItemHandle, CacheQuality, CacheTimestamp);
  CanonicalDataType:= VarType(CacheValue);
  AnalogType:= not VarIsArray(CacheValue) and
                NumericType(CanonicalDataType)
end;

procedure TServerItemRef.GetMaxAgeVQT(MaxAge: Cardinal;
  ActualTimestamp: TFileTime; var Result: OleVariant; var Quality: Word;
  var Timestamp: TFileTime);
begin
  // Read new value
  if MaxAge = 0 then
    GetItemVQT(Result, Quality, Timestamp)
  // Read cached value
  else if MaxAge = MAXDWORD then
    GetCacheVQT(Result, Quality, Timestamp)
  else begin
    GetCacheVQT(Result, Quality, Timestamp);
    // Check MaxAge
    if Int64(Timestamp) + Int64(MaxAge) * 10000 < Int64(ActualTimestamp) then
      GetItemVQT(Result, Quality, Timestamp);
  end;
end;

procedure TServerItemRef.GetCacheVQT(var Result: OleVariant;
  var Quality: Word; var Timestamp: TFileTime);
begin
  if (AccessRights and OPC_READABLE) = 0 then
    raise EOpcError.Create(OPC_E_BADRIGHTS);
  Result:= CacheValue;
  Quality:= CacheQuality;
  Timestamp:= CacheTimestamp;
end;

procedure TServerItemRef.GetItemVQT(var Result: OleVariant;
  var Quality: Word; var Timestamp: TFileTime);
begin
  if (AccessRights and OPC_READABLE) = 0 then
    raise EOpcError.Create(OPC_E_BADRIGHTS);
  Quality:= OPC_QUALITY_GOOD;
  Result:= GOpcItemServer.GetItemVQT(ItemHandle, Quality, Timestamp);

  CacheValue:= Result;
  CacheQuality:= Quality;
  CacheTimestamp:= Timestamp;
end;

procedure TServerItemRef.ReleaseNonGroupReference;
begin
  Destroy
end;

procedure TServerItemRef.SetItemVQT(const ValueVQT: OPCITEMVQT);
var
  TypeValVQT: OPCITEMVQT;
  Res: HRESULT;
begin
  if (AccessRights and OPC_WRITABLE) = 0 then  {cf 1.13.15}
    raise EOpcError.Create(OPC_E_BADRIGHTS);
  if VarType(ValueVQT.vDataValue) = VT_EMPTY then
    raise EOpcError.Create(OPC_E_BADTYPE);
  try
    if VarType(ValueVQT.vDataValue) = CanonicalDataType then
    begin
      GOpcItemServer.SetItemVQT(ItemHandle, ValueVQT)
    end else
    begin
      TypeValVQT := ValueVQT;
      Res:= VariantChangeType(TypeValVQT.vDataValue, ValueVQT.vDataValue, 0, CanonicalDataType);
      if not Succeeded(Res) then
        raise EOpcError.Create(Res);
      GOpcItemServer.SetItemVQT(ItemHandle, TypeValVQT)
    end
  except
    on E: EVariantError do
      raise EOpcError.Create(OPC_E_BADTYPE)
  end
end;

procedure TServerItemRef.SetItemValue(const Value: OleVariant);
var
  TypeVal: OleVariant;
  Res: HRESULT;
begin
  if (AccessRights and OPC_WRITABLE) = 0 then  {cf 1.13.15}
    raise EOpcError.Create(OPC_E_BADRIGHTS);
  if VarType(Value) = VT_EMPTY then
    raise EOpcError.Create(OPC_E_BADTYPE);
  try
    if VarType(Value) = CanonicalDataType then
    begin
      GOpcItemServer.SetItemValue(ItemHandle, Value)
    end else
    begin
      Res:= VariantChangeType(TypeVal, Value, 0, CanonicalDataType);
      if not Succeeded(Res) then
        raise EOpcError.Create(Res);
      GOpcItemServer.SetItemValue(ItemHandle, TypeVal)
    end
  except
    on E: EVariantError do
      raise EOpcError.Create(OPC_E_BADTYPE)
  end
end;

destructor TServerItemRef.Destroy;

begin
  if ItemHandle <> INVALID_HANDLE_VALUE then
    GOpcItemServer.ReleaseHandle(ItemHandle);
  inherited Destroy
end;

{ TServerItem }

procedure TServerItem.AddRef(GroupItem: TGroupItemImpl);
begin
  with LockRefList do
  try
    Add(GroupItem)
  finally
    UnlockRefList
  end
end;

constructor TServerItem.Create(const aItemID: String;
                       aOwner: TServerItemList;
                       aGroupItem: TGroupItemImpl); {first to create}
begin
  inherited Create(aItemID);
  Owner:= aOwner;
  RefList:= TList.Create;
  AddRef(aGroupItem);
  Subscribed:= GOpcItemServer.SubscribeToItem(ItemHandle, ItemCallback)
end;

destructor TServerItem.Destroy;
begin
  {this is usually called from inside a RefList lock}
  if Subscribed then
    GOpcItemServer.UnsubscribeToItem(ItemHandle);
  if Assigned(Owner) then
    Owner.DeleteItem(ItemID);
  RefList.Free;
  inherited Destroy
end;

class function TServerItemRef.GetItem(const aItemID: String): TServerItemRef;
var
  i: Integer;
begin
    {if exists in list then return new TServerItemRef, else return
    pointer to item in the list}
  with ServerItemList do
  begin
    with LockList do
    try
      i:= IndexOf(aItemID);
      if i = -1 then
        Result:= TServerItemRef.Create(aItemId)
      else
        Result:= TServerItemRef(Objects[i])
    finally
      UnlockList
    end
  end
end;

procedure TServerItem.ItemCallback(const Value: OleVariant; Quality: Word; TimeStamp: TFileTime);
var
  i: Integer;
begin
  with LockRefList do
  try
    CacheValue := Value;
    CacheQuality := Quality;
    if Int64(TimeStamp) = Int64(TimestampNotSet) then
      GOpcItemServer.GetTimestamp(Timestamp);
    CacheTimestamp := TimeStamp;
    for i:= 0 to Count - 1 do
      TGroupItemImpl(Items[i]).ItemCallback(Value, Quality, Timestamp)
  finally
    UnlockRefList
  end
end;

function TServerItem.LockRefList: TList;
begin
  EnterCriticalSection(Owner.FRefListCritSect);
  Result:= RefList
end;

procedure TServerItem.ReleaseNonGroupReference;
begin
end;

procedure TServerItem.ReleaseRef(GroupItem: TGroupItemImpl);
var
  i: Integer;
  DoDestroy: Boolean;
begin
  with LockRefList do
  try
    i:= RefList.IndexOf(GroupItem);
    if i <> -1 then
      RefList.Delete(i);
    DoDestroy:= RefList.Count = 0;
  finally
    UnlockRefList
  end;
  if DoDestroy then
    Destroy
end;

procedure TServerItem.UnlockRefList;
begin
  LeaveCriticalSection(Owner.FRefListCritSect)
end;


{ TServerItemList }

function TServerItemList.AddServerItem(GroupItem: TGroupItemImpl; const ItemID: string): TServerItem;
var
  i: Integer;
begin
  with LockList do
  try
{$IFDEF Evaluation}
    if Count > MaxEvalItems then
    begin
      Result:= nil; {why?}
      raise EOpcError.Create(OPC_E_RANGE)
    end;
{$ENDIF}
    i:= IndexOf(ItemID);
    if i = -1 then
    begin
      Result:= TServerItem.Create(ItemID, Self, GroupItem);
      AddObject(ItemID, Result)
    end else
    begin
      Result:= TServerItem(Objects[i]);
      Result.AddRef(GroupItem)
    end
  finally
    UnlockList
  end
end;

constructor TServerItemList.Create;
begin
  inherited Create;
  FList:= TStringList.Create;
  InitializeCriticalSection(FCritSect);
  InitializeCriticalSection(FRefListCritSect);
  FList.Sorted:= true;
  FList.Duplicates:= dupError
end;

procedure TServerItemList.DeleteItem(const ItemID: String);
var
  i: Integer;
begin
  with LockList do
  try
    if Find(ItemID, i) then
      Delete(i)
  finally
    UnlockList
  end
end;

destructor TServerItemList.Destroy;
begin
  EnterCriticalSection(FCritSect);
  with FList do
  try
    while Count > 0 do
      TServerItem(Objects[0]).Free;
    Free
  finally
    LeaveCriticalSection(FCritSect)
  end;
  DeleteCriticalSection(FCritSect);
  DeleteCriticalSection(FRefListCritSect);
  inherited Destroy
end;

function TServerItemList.LockList: TStringList;
begin
  EnterCriticalSection(FCritSect);
  Result:= FList
end;

procedure TServerItemList.UnlockList;
begin
  LeaveCriticalSection(FCritSect)
end;

{ TGroupItemInfo }

function TGroupItemInfo.GetItemID: String;
begin
  Result:= TServerItem(GetServerItem).ItemID
end;

function TGroupItemInfo.GetItemHandle: TItemHandle;
begin
  Result:= TServerItem(GetServerItem).ItemHandle
end;

function TGroupItemInfo.GetCanonicalDataType: TVarType;
begin
  Result:= TServerItem(GetServerItem).CanonicalDataType
end;

function TGroupItemInfo.GetAccessRights: TAccessRights;
begin
  Result:= NativeAccessRights(TServerItem(GetServerItem).AccessRights)
end;

function TGroupItemInfo.ItemEUInfo: IEUInfo;
begin
  Result:= TServerItem(GetServerItem).EUInfo
end;

function TGroupItemInfo.ItemProperties: IItemProperties;
begin
  Result:= TServerItem(GetServerItem).ItemProperties
end;



{ TGroupItemImpl }

procedure TGroupItemImpl.AssignTo(var Result: OPCITEMATTRIBUTES);
begin
  FillChar(Result, SizeOf(Result), 0);
  ServerItem.AssignTo(Result);
  if FActive then  {cf 1.01.7}
    DWORD(Result.bActive):= 1
  else
    FActive:= false;
  Result.hClient:= FhClient;
  Result.hServer:= Integer(Self);
  Result.vtRequestedDataType:= FRequestedDataType
end;

constructor TGroupItemImpl.Create(aOwner: TGroupImpl; const ItemDef: OPCITEMDEF;
  var ItemResult: OPCITEMRESULT);
var
  ItemID: String;
begin
  inherited Create;
  FillChar(ItemResult, SizeOf(ItemResult), 0);
  InitItemID(ItemDef, ItemID);
  ServerItem:= ServerItemList.AddServerItem(Self, ItemID);
  Owner:= aOwner;
  FhClient:= ItemDef.hClient;
  CheckRequestedVarType(ItemDef.vtRequestedDataType);
  FRequestedDataType:= ItemDef.vtRequestedDataType;
  SetActive(ItemDef.bActive);
  {get results}
  with ItemResult do
  begin
    hServer:= OPCHANDLE(Self);
    vtCanonicalDataType:= ServerItem.CanonicalDataType;
    dwAccessRights:= ServerItem.AccessRights
  end;
  {notify in main thread}
  SendMessage(OpcWindow, CM_NOTIFICATION, nfAddItem, Integer(Self));
  Include(GroupItemState, gisNotified) {cf 1.14.23}
end;

constructor TGroupItemImpl.CreateClone(aOwner: TGroupImpl; Source: TGroupItemImpl);
begin
  inherited Create;
  Owner:= aOwner;
  FhClient:= Source.FhClient;
  FRequestedDataType:= Source.FRequestedDataType;
  ServerItem:= Source.ServerItem;
  ServerItem.AddRef(Self);
  SetActive(Source.Active)  {cf 1.01.13}
end;

destructor TGroupItemImpl.Destroy;
begin
  {notify in main thread}
  if gisNotified in GroupItemState then  {cf 1.14.23}
  begin
    SendMessage(OpcWindow, CM_NOTIFICATION, nfRemoveItem, Integer(Self));
    Exclude(GroupItemState, gisNotified)
  end;
  if Assigned(ServerItem) then
    ServerItem.ReleaseRef(Self);
  inherited Destroy
end;

procedure TGroupItemImpl.GetCacheValue(var Result: OleVariant;
  var Quality: Word; var Timestamp: TFiletime);
var
  Res: HRESULT;
begin
  Result:= CacheValue;
  Quality:= CacheQuality;
  Timestamp:= CacheTimestamp;
  if (FRequestedDataType <> VT_EMPTY) and
     (FRequestedDataType <> ServerItem.CanonicalDataType) then {1.10.3}
  begin
    Res:= VariantChangeType(Result, Result, 0, FRequestedDataType);
    if not Succeeded(Res) then
      raise EOpcError.Create(Res)
  end
end;

procedure TGroupItemImpl.GetMaxAgeValue(MaxAge: Cardinal; ActualTimestamp: TFileTime;
  var Result: OleVariant; var Quality: Word; var Timestamp: TFiletime);
begin
  // Read new value
  if MaxAge = 0 then
    GetItemValue(Result, Quality, Timestamp)
  // Read cached value
  else if MaxAge = MAXDWORD then
    GetCacheValue(Result, Quality, Timestamp)
  else begin
    GetCacheValue(Result, Quality, Timestamp);
    // Check MaxAge
    if Int64(Timestamp) + Int64(MaxAge) * 10000 < Int64(ActualTimestamp) then
      GetItemValue(Result, Quality, Timestamp);
  end;
end;

procedure TGroupItemImpl.GetItemValue(var Result: OleVariant;
  var Quality: Word; var Timestamp: TFiletime);
var
  Res: HRESULT;
begin
  ServerItem.GetItemVQT(Result, Quality, Timestamp);
  CacheValue:= Result;
  CacheQuality:= Quality;
  CacheTimestamp:= Timestamp;
  if (FRequestedDataType <> VT_EMPTY) and
     (FRequestedDataType <> ServerItem.CanonicalDataType) then {1.10.3}
  begin
    Res:= VariantChangeType(Result, Result, 0, FRequestedDataType);
    if not Succeeded(Res) then
      raise EOpcError.Create(Res)
  end
end;

procedure TGroupItemImpl.InvalidateCache;
begin
  ServerItem.GetItemVQT(CacheValue, CacheQuality, CacheTimestamp);
  CacheUpdated:= true
end;

procedure TGroupItemImpl.ItemCallback(const Value: OleVariant; Quality: Word; Timestamp: TFiletime);
begin
  if UpdateCache(Value, Quality, Timestamp) then
    CacheUpdated:= true
end;

function TGroupItemImpl.GetServerItem: TObject;
begin
  Result:= ServerItem
end;

procedure TGroupItemImpl.SetActive(Value: Boolean);
begin
  if FActive <> Value then
  begin
    FActive:= Value;
    if Value then
      InvalidateCache
  end
end;

function TGroupItemImpl.Tick: Boolean;
var
  ItemVal: OleVariant;
  ItemQuality: Word;
  ItemTimestamp: TFileTime;
begin
  if not FActive then
  begin
    Result:= false
  end else
  if CacheUpdated then
  begin
    Result:= true;
    CacheUpdated:= false
  end else
  if not ServerItem.Subscribed then
  begin
    ServerItem.GetItemVQT(ItemVal, ItemQuality, ItemTimestamp);
    Result:= UpdateCache(ItemVal, ItemQuality, ItemTimestamp)
  end else
  begin
    Result:= false
  end
end;

function TGroupItemImpl.UpdateCache(const Value: OleVariant;
  Quality: Word; Timestamp: TFileTime): Boolean;

  function ValueChanged: Boolean;  {cf 1.14.17}
  var
    EUI: IEUInfo;
    Limits: OleVariant;
    ah, al: Double;
  begin
    EUI:= ItemEUInfo;
    if Assigned(EUI) and
       TServerItem(GetServerItem).AnalogType and
       (EUI.EUType = euAnalog) then
    begin
      try
        Limits:= EUI.EUInfo;
        al:= Limits[0];
        ah:= Limits[1]
      except
        on EVariantError do  {the limits are no good. Ignore deadband &&&}
        begin
          al:= 0;
          ah:= 0
        end
      end;
      if (    PercentDeadbandSet and (Abs(CacheValue - Value) > (      PercentDeadband/100.0)*(ah - al))) or
         (not PercentDeadbandSet and (Abs(CacheValue - Value) > (Group.PercentDeadband/100.0)*(ah - al))) then
      begin
        Result:= true
      end else
      begin
        Result:= false;
        {timestamp must be updated. This seems odd, but is required by
         Section 4.5.1.6 of DA spec}
        CacheTimestamp := Timestamp;
      end
    end else
    begin
      Result:= not CompareVariant(CacheValue, Value)
    end
  end;

begin
  Result:= (Quality <> CacheQuality) or
           ValueChanged;
  if Result then  {cf 1.13.14}
  begin
    CacheValue:= Value;
    CacheQuality:= Quality;
    CacheTimestamp:= Timestamp;
    GOpcItemServer.OnItemValueChange(Self)
  end
end;

function TGroupItemImpl.LastUpdateTime: TDateTime;
var
  st: TSystemTime;
begin
  FileTimeToSystemTime(CacheTimestamp, st);
  Result:= SystemTimeToDateTime(st)
end;

function TGroupItemImpl.LastUpdateValue: OleVariant;
begin
  Result:= CacheValue
end;

function TGroupItemImpl.Group: TGroupInfo;
begin
  Result:= Owner
end;

{ TGroupItemList }

function TGroupItemList.GetGroupItem(i: Integer): TGroupItemImpl;
begin
  Result:= TGroupItemImpl(Items[i])  {cf 1.14.17}
end;

{ TGroupList }

procedure TGroupList.AddGroup(const Name: String; aGroup: TGroupImpl);
begin
  if FindGroup(Name) then
    raise EOpcError.Create(OPC_E_DUPLICATENAME);
  AddObject(Name, aGroup)
end;

constructor TGroupList.Create;
begin
  inherited Create;
  Sorted:= true;
  Duplicates:= dupError
end;

destructor TGroupList.Destroy;
var
  i: Integer;
begin
  for i:= 0 to Count - 1 do
    Group(i).Free;
  inherited Destroy
end;

function TGroupList.FindGroup(const Name: String): Boolean;
var
  i: Integer;
begin
  Result:= Find(Name, i)
end;

function TGroupList.GetUniqueGroupName: String;
var
  i, j: Integer;
begin
  i:= 0;
  repeat
    Result:= 'Group' + IntToStr(i);
    Inc(i)
  until not Find(Result, j)
end;

function TGroupList.Group(i: Integer): TGroupImpl;
begin
  Result:= TGroupImpl(Objects[i])
end;

procedure TGroupList.GroupByName(const Name: String; const riid: TIID; out ppUnk: IUnknown);
var
  i: Integer;
begin
  if not Find(Name, i) then
    raise EOpcError.Create(E_INVALIDARG);
  if not Group(i).GetInterface(riid, ppUnk) then
    raise EOpcError.Create(E_NOINTERFACE)
end;

procedure TGroupList.UpdateNames;
var
  i : Integer;
begin
  Sorted := False;
  for i := 0 to Count-1 do
   Strings[i] := TGroupImpl(Objects[i]).Name;
  Sorted := True;
end;

{ TAsyncTask }

constructor TAsyncTask.Create(aGroup: TGroupImpl; aTransactionID: DWORD);
begin
  inherited Create;
  Group:= aGroup;
  TransactionID:= aTransactionID;
  Group.FTaskList.Add(Self)
{$IFDEF GLD}
  ; SendDebug(Format('%s.Create %p. Transid=%d', [ClassName, Pointer(Self), TransactionId]));
{$ENDIF}
end;

procedure TAsyncTask.DoProcess; {cf 1.15.5.2}
begin
  if not Deleted then
  begin
{$IFDEF GLD}
    SendDebug(Format('AsyncTask.Process %p, Transid=%d', [Pointer(Self), TransactionId]));
{$ENDIF}
    Group.DeleteTask(Self);  {cf 1.14.12}
    if not BlockProcess then
      Process
  end;
end;

destructor TAsyncTask.Destroy;
begin
{$IFDEF GLD}
  SendDebug(Format('AsyncTask.Destroy %p', [Pointer(Self)]));
{$ENDIF}
  inherited Destroy
end;

{ TAsyncIOTask }

constructor TAsyncIOTask.Create(aGroup: TGroupImpl; aTransactionID, aCount: DWORD;
  phServer: POPCHANDLEARRAY; ppErrors: PResultList; var Result: HResult);
var
  i: Integer;
begin
  inherited Create(aGroup, aTransactionID);
  Result := S_OK;
  Count:= aCount;
  SetLength(GroupItem, Count);
  SetLength(GroupItemValid, Count);
  TransactionID:= aTransactionID;
  ValidCount := 0;
  for i:= 0 to Count - 1 do
  begin
    GroupItem[i]:= TGroupItemImpl(phServer^[i]);
    try
      Group.ValidateServerHandle(Cardinal(GroupItem[i]));
      GroupItemValid[i] := True;
      Inc(ValidCount);
    except
      on E: EOPCError do
      begin
        ppErrors[i] := E.ErrorCode;
        GroupItemValid[i] := False;
        Result := S_FALSE;
      end;
    end;
  end;

  if ValidCount = 0 then
  begin
    BlockProcess := True;
    Result := S_FALSE;
  end;
end;

function TAsyncIOTask.ProcessItems(ppErrors: PResultList; var Data): HRESULT;
var
  i, Index: Integer;
  IsError: Boolean;
begin
  Result := S_OK;
  try
    IsError:= False;
    Index := 0;
    for i:= 0 to Count - 1 do
    if GroupItemValid[i] then
    begin
      try
        Group.ValidateServerHandle(Cardinal(GroupItem[i]));
        ProcessItem(TGroupItemImpl(GroupItem[i]), i, Index, Data, Result);
        if ppErrors <> nil then
          ppErrors^[Index]:= S_OK;
      except
        on E: EOpcError do
        begin
          IsError:= True;
          if ppErrors <> nil then
            ppErrors^[Index]:= E.ErrorCode
        end
      end;
      Inc(Index);
    end;
    if IsError then
      Result:= S_FALSE
  except
    on E: EOpcError do
    begin
      Result:= E.ErrorCode;
    end
  end
end;

{ TAsyncWriteTask }

constructor TAsyncWriteTask.Create(aGroup: TGroupImpl; aTransactionID, aCount: DWORD;
  phServer: POPCHANDLEARRAY; pValues: POleVariantArray;
  ppErrors: PResultList; var Result: HResult);
var
  i: Integer;
begin
  inherited Create(aGroup, aTransactionID, aCount, phServer, ppErrors, Result);
  SetLength(Values, Count);
  for i:= 0 to Count - 1 do
    Values[i]:= pValues^[i];
end;

procedure TAsyncWriteTask.Process;
var
  hClientItem: array of OPCHANDLE;
  Errors: array of HRESULT;
  MasterError: HRESULT;
begin
  SetLength(hClientItem, ValidCount);
  SetLength(Errors, ValidCount);
  MasterError:= ProcessItems(Pointer(Errors), hClientItem);
  with Group do
  if Assigned(FDataCallback) then
  begin
    try   {try except cf 1.14.4}
      FDataCallback.OnWriteComplete(TransactionID,
          hClientGroup, MasterError, ValidCount,
          Pointer(hClientItem), Pointer(Errors))
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Group, ccDa2WriteComplete)
    end
  end
end;

procedure TAsyncWriteTask.ProcessItem(Item: TGroupItemImpl; IndexSrc,
  IndexDst: Integer; var Data; var Status: HRESULT);
var
  pClienthandles: POPCHANDLEARRAY absolute Data;
begin
  pClienthandles[IndexDst]:= Item.FhClient;
  Item.ServerItem.SetItemValue(Values[IndexSrc]);
end;

{ TAsyncWriteVQTTask }

constructor TAsyncWriteVQTTask.Create(aGroup: TGroupImpl; aTransactionID,
  aCount: DWORD; phServer: POPCHANDLEARRAY; pValues: POPCITEMVQTARRAY;
  ppErrors: PResultList; var Result: HResult);
var
  i: Integer;
begin
  inherited Create(aGroup, aTransactionID, aCount, phServer, ppErrors, Result);
  SetLength(Values, Count);
  for i:= 0 to Count - 1 do
    Values[i]:= pValues^[i];
end;

procedure TAsyncWriteVQTTask.Process;
var
  hClientItem: array of OPCHANDLE;
  Errors: array of HRESULT;
  MasterError: HRESULT;
begin
  SetLength(hClientItem, ValidCount);
  SetLength(Errors, ValidCount);
  MasterError:= ProcessItems(Pointer(Errors), hClientItem);
  with Group do
  if Assigned(FDataCallback) then
  begin
    try   {try except cf 1.14.4}
      FDataCallback.OnWriteComplete(TransactionID,
          hClientGroup, MasterError, ValidCount,
          Pointer(hClientItem), Pointer(Errors))
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Group, ccDa2WriteComplete)
    end
  end
end;

procedure TAsyncWriteVQTTask.ProcessItem(Item: TGroupItemImpl;
  IndexSrc, IndexDst: Integer; var Data; var Status: HRESULT); 
var
  pClienthandles: POPCHANDLEARRAY absolute Data;
begin
  pClienthandles[IndexDst]:= Item.FhClient;
  Item.ServerItem.SetItemVQT(Values[IndexSrc]);
end;

{ TAsyncReadTask }

procedure TAsyncReadTask.Process;
var
  Errors: array of HRESULT;
  IOP: TAsyncReadParam;
  MasterError: HRESULT;
begin
{$IFDEF GLD}
  SendDebug(Format('AsyncIO2Read.Process dwCount = %d', [Count]));
{$ENDIF}
  SetLength(IOP.hClientItem, ValidCount);
  SetLength(Errors, ValidCount);
  SetLength(IOP.Qualities, ValidCount);
  SetLength(IOP.Timestamps, ValidCount);
  SetLength(IOP.Values, ValidCount);
  MasterError:= ProcessItems(Pointer(Errors), IOP);
  with Group do
  if Assigned(FDataCallback) then
  begin
    try   {try except cf 1.14.5}
      FDataCallback.OnReadComplete(TransactionID,
       hClientGroup, S_OK, MasterError, ValidCount,
       Pointer(IOP.hClientItem), Pointer(IOP.Values), Pointer(IOP.Qualities),
       Pointer(IOP.Timestamps), Pointer(Errors))
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Group, ccDa2ReadComplete)
    end
  end;
end;

procedure TAsyncReadTask.ProcessItem(Item: TGroupItemImpl; IndexSrc,
  IndexDst: Integer; var Data; var Status: HRESULT);
var
  IOP: TAsyncReadParam absolute Data;
begin
  with GroupItem[IndexSrc] do
  begin
    IOP.hClientItem[IndexDst]:= FhClient;
    GetItemValue(IOP.Values[IndexDst], IOP.Qualities[IndexDst], IOP.Timestamps[IndexDst]);
  end;
end;

{ TAsyncReadMaxAgeTask }

constructor TAsyncReadMaxAgeTask.Create(aGroup: TGroupImpl; aTransactionID,
  aCount: DWORD; phServer: POPCHANDLEARRAY; pdwMaxAge: PDWORDARRAY;
  ppErrors: PResultList; var Result: HResult);
var
  I : Integer;
begin
  inherited Create(aGroup, aTransactionID, aCount, phServer, ppErrors, Result);
  SetLength(MaxAge, Count);
  for I:= 0 to Count - 1 do
    MaxAge[I]:= pdwMaxAge[I];
  GOpcItemServer.GetTimestamp(ActualTimestamp);
end;

procedure TAsyncReadMaxAgeTask.Process;
var
  Errors: array of HRESULT;
  IOP: TAsyncReadMaxAgeParam;
  MasterError: HRESULT;
begin
{$IFDEF GLD}
  SendDebug(Format('AsyncIO3ReadMaxAge.Process dwCount = %d', [Count]));
{$ENDIF}
  SetLength(IOP.hClientItem, ValidCount);
  SetLength(Errors, ValidCount);
  SetLength(IOP.Qualities, ValidCount);
  SetLength(IOP.Timestamps, ValidCount);
  SetLength(IOP.Values, ValidCount);
  MasterError:= ProcessItems(Pointer(Errors), IOP);
  with Group do
  if Assigned(FDataCallback) then
  begin
    try   {try except cf 1.14.5}
      FDataCallback.OnReadComplete(TransactionID,
       hClientGroup, S_OK, MasterError, ValidCount,
       Pointer(IOP.hClientItem), Pointer(IOP.Values), Pointer(IOP.Qualities),
       Pointer(IOP.Timestamps), Pointer(Errors))
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Group, ccDa2ReadComplete)
    end
  end
end;

procedure TAsyncReadMaxAgeTask.ProcessItem(Item: TGroupItemImpl; IndexSrc,
  IndexDst: Integer; var Data; var Status: HRESULT);
var
  IOP: TAsyncReadMaxAgeParam absolute Data;
begin
  IOP.hClientItem[IndexDst]:= Item.FhClient;
  Item.GetMaxAgeValue(MaxAge[IndexSrc], ActualTimestamp,
    IOP.Values[IndexDst], IOP.Qualities[IndexDst], IOP.Timestamps[IndexDst]);
end;

{ TRefreshTask }

constructor TRefreshTask.Create(aGroup: TGroupImpl; aTransactionID: DWORD;
  aSource: OPCDATASOURCE);
var
  i: Integer;
begin
  inherited Create(aGroup, aTransactionID);
  Source:= aSource;
  ActiveList:= TList.Create;
  with Group do
  for i:= 0 to FItemList.Count - 1 do
    if FItemList[i].FActive then
      ActiveList.Add(FItemList[i]);
  if ActiveList.Count = 0 then
  begin
    Group.DeleteTask(Self);
    Deleted:= true;        {cf 14.}
    raise EOpcError.Create(E_FAIL)
  end
end;

destructor TRefreshTask.Destroy; {cf 1.15.3}
begin
  ActiveList.Free;
  inherited Destroy
end;

procedure TRefreshTask.Process;
begin
  Group.DoRefresh2(Source, TransactionID, ActiveList)
end;

{ TRefreshMaxAgeTask }

constructor TRefreshMaxAgeTask.Create(aGroup: TGroupImpl; aTransactionID: DWORD;
  aMaxAge: DWORD);
var
  i: Integer;
begin
  inherited Create(aGroup, aTransactionID);
  MaxAge:= aMaxAge;
  GOpcItemServer.GetTimestamp(ActualTimestamp);
  ActiveList:= TList.Create;
  with Group do
  for i:= 0 to FItemList.Count - 1 do
    if FItemList[i].FActive then
    begin
      ActiveList.Add(FItemList[i]);
    end;
  if ActiveList.Count = 0 then
  begin
    Group.DeleteTask(Self);
    Deleted:= true;
    raise EOpcError.Create(E_FAIL)
  end
end;

destructor TRefreshMaxAgeTask.Destroy;
begin
  ActiveList.Free;
  inherited Destroy
end;

procedure TRefreshMaxAgeTask.Process;
begin
  Group.DoRefreshMaxAge(MaxAge, ActualTimestamp, TransactionID, ActiveList)
end;

{ TRefreshTask1 }

constructor TRefreshTask1.Create(aGroup: TGroupImpl; aSource: OPCDATASOURCE;
  aFormat: TDa1Format);
begin
  inherited Create(aGroup, 0, aSource);
  Format:= aFormat
end;

procedure TRefreshTask1.Process;
begin
  Group.DoRefresh1(Source, Integer(Self), Format, ActiveList)
end;

{ TAsyncWriteTask1 }

constructor TAsyncWriteTask1.Create(aGroup: TGroupImpl; aCount: DWORD;
  phServer: POPCHANDLEARRAY; pValues: POleVariantArray;
  ppErrors: PResultList; var Result: HResult);
begin
  inherited Create(aGroup, 0, aCount, phServer, pValues, ppErrors, Result)
end;

procedure TAsyncWriteTask1.Process;
var
  GHW: POPCGROUPHEADERWRITE;
  IHW: POPCITEMHEADERWRITEARRAY;
  stgmed: TStgMedium;
  Fe: TFormatEtc;
  StreamSize: Integer;
begin
  with Fe do
  begin
    cfFormat:= GDa1Format[da1WriteComplete];
    ptd:= nil;
    dwAspect:= DVASPECT_CONTENT;
    lindex:= -1;
    tymed:= TYMED_HGLOBAL
  end;
  stgmed.tymed:= TYMED_HGLOBAL;
  StreamSize:= SizeOf(OPCGROUPHEADERWRITE) + ValidCount * SizeOf(OPCITEMHEADERWRITE);
  stgmed.hGlobal:= GlobalAlloc(GMEM_FIXED, StreamSize);
  stgmed.UnkForRelease:= nil;
  try
    GHW:= POPCGROUPHEADERWRITE(stgmed.hGlobal);
    with GHW^ do
    begin
      dwItemCount:= ValidCount;
      hClientGroup:= Group.hClientGroup;
      dwTransactionID:= TransactionID;
      hrStatus:= S_OK
    end;
    IHW:= POPCITEMHEADERWRITEARRAY(stgmed.hGlobal + SizeOf(OPCGROUPHEADERWRITE));
    ProcessItems(nil, IHW);
    try
      with Group do
        if Assigned(FDa1Advise[da1WriteComplete]) then  {cf 1.14.7}
          FDa1Advise[da1WriteComplete].OnDataChange(fe, stgmed)
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Group, ccDa1WriteComplete)
    end
  finally
    GlobalFree(StgMed.hGlobal)
  end;
end;

procedure TAsyncWriteTask1.ProcessItem(Item: TGroupItemImpl; IndexSrc,
  IndexDst: Integer; var Data; var Status: HRESULT);
var
  IHW: POPCITEMHEADERWRITEARRAY absolute Data;
begin
  with IHW^[IndexDst] do
  begin
    try
      hClient:= Item.FhClient;
      Item.ServerItem.SetItemValue(Values[IndexSrc]);
      dwError:= S_OK
    except
      on E: EOpcError do
        dwError:= E.ErrorCode
    end
  end;
end;

{ TAsyncReadTask1 }

constructor TAsyncReadTask1.Create(aGroup: TGroupImpl;
  aCount: DWORD; phServer: POPCHANDLEARRAY; aSource: OPCDATASOURCE;
  aFormat: TDa1Format; ppErrors: PResultList; var Result: HResult);
begin
  inherited Create(aGroup, 0, aCount, phServer, ppErrors, Result);
  Format:= aFormat;
  Source:= aSource
end;

procedure TAsyncReadTask1.Process;
var
  stgmed: TStgMedium;
  Fe: TFormatEtc;
  Stream: TDataChangeStream;
  Status: HRESULT;
begin
  Stream:= CreateStream;
  try
    with Fe do
    begin
      cfFormat:= GDa1Format[Stream.Format];
      ptd:= nil;
      dwAspect:= DVASPECT_CONTENT;
      lindex:= -1;
      tymed:= TYMED_HGLOBAL
    end;
    with Stream.GroupHeader^ do
    begin
      dwItemCount:= ValidCount;
      hClientGroup:= Group.hClientGroup;
      dwTransactionID:= DWORD(Self)
    end;
    Status:= ProcessItems(nil, Stream);
    with Stream.GroupHeader^ do
    begin
      dwSize:= Stream.Position;
      hrStatus:= Status
    end;
    stgmed.tymed:= TYMED_HGLOBAL;
    stgmed.hGlobal:= GlobalHandle(Stream.Memory);
    try   {try except cf 1.14.8}
      with Group do
        if Assigned(FDa1Advise[Stream.Format]) then
          FDa1Advise[Stream.Format].OnDataChange(fe, stgmed)
    except
      on E: EOleSysError do
        GOpcItemServer.GroupCallbackError(E, Group, ccDa1ReadComplete)
    end
  finally
    Stream.Free
  end
end;

function TAsyncReadTask1.CreateStream: TDataChangeStream;
begin
  if Format = Da1DataTime then
    Result:= TDataChangeStream1.Create(ValidCount)
  else
    Result:= TDataChangeStream2.Create(ValidCount)
end;

procedure TAsyncReadTask1.ProcessItem(Item: TGroupItemImpl; IndexSrc,
  IndexDst: Integer; var Data; var Status: HRESULT);
var
  Stream: TDataChangeStream absolute Data;
  Value: OleVariant;
  Timestamp: TFiletime;
begin
   with Stream.ItemHeader2(IndexSrc)^ do
   begin
     wReserved:= 0;
     dwValueOffset:= Stream.Position;
     Value:= Null;
     hClient:= 0;
     try
       hClient:= Item.FhClient;
       if Source = OPC_DS_DEVICE then
         Item.GetItemValue(Value, wQuality, Timestamp)
       else
         Item.GetCacheValue(Value, wQuality, Timestamp);
       Stream.SetTimestamp(IndexDst, Timestamp)
     except
       on EOpcError do
         wQuality:= OPC_QUALITY_BAD
     end;
     if (Status = S_OK) and (wQuality <> OPC_QUALITY_GOOD) then
       Status:= S_FALSE
   end;
   {note that any writes to the stream may alter the
   base pointer, therefore we MUST NOT write to the
   stream inside with Stream.ItemHeader1(i)^ do}
   Stream.WriteVariant(Value)  {cf 1.13.12}
end;

{ TDataChangeStream }

constructor TDataChangeStream.Create(aCount: Integer);
var
  HeaderSize: Integer;
begin
  inherited Create;
  HeaderSize:= SizeOf(OPCGROUPHEADER) + aCount*ItemHeaderSize;
  SetSize(HeaderSize);
  Seek(HeaderSize, soFromBeginning)
end;

procedure TDataChangeStream.SetTimestamp(i: Integer;
  const aTimestamp: TFiletime);
begin
end;

function TDataChangeStream.GroupHeader: POPCGROUPHEADER;
begin
  Result:= POPCGROUPHEADER(Memory)
end;

procedure TDataChangeStream.WriteVariant(const Value: OleVariant);

procedure WriteString(const StringVal: WideString); {cf 1.13.14}
{in a 'Noncompliant' DA1 stream the byte count includes the null.
 Versions of the toolkit prior to release 1.13 produced these
 streams. They seemed to work OK}
const
  NullChar: WideChar = #0;
var
  StrLen: Integer;
  NoncompliantDA1Stream: Boolean;
begin
  NoncompliantDA1Stream:= GOpcItemServer.IncludeNullInDA1ByteCount;
  if NoncompliantDA1Stream then
    StrLen:= (Length(StringVal) + 1) * SizeOf(WideChar)
  else
    StrLen:= Length(StringVal) * SizeOf(WideChar);
  WriteBuffer(StrLen, SizeOf(StrLen));
  if not NoncompliantDA1Stream then
    Inc(StrLen, 2);
  if StringVal = '' then
    WriteBuffer(NullChar, SizeOf(NullChar))
  else
    WriteBuffer(Pointer(StringVal)^, StrLen)
end;

{ TVarArrayBound = packed record
    ElementCount: Integer;
    LowBound: Integer;
  end;
  TVarArrayBoundArray = array [0..0] of TVarArrayBound;
  PVarArrayBoundArray = ^TVarArrayBoundArray;
  TVarArrayCoorArray = array [0..0] of Integer;
  PVarArrayCoorArray = ^TVarArrayCoorArray;

  PVarArray = ^TVarArray;
  TVarArray = packed record
    DimCount: Word;
    Flags: Word;
    ElementSize: Integer;
    LockCount: Integer;
    Data: Pointer;
    Bounds: TVarArrayBoundArray;
  end;}
{ref: System.pas}

const
  {note D5 and D6 have different definitions of TVarArray.
   SizeOf(TVarArray) will not work declaration should work for either}
  SafeArraySize = SizeOf(TSafeArray);

var
  VD: TVarData absolute Value;
  VarArray: TVarArray;
  i: Integer;
  Ptr: Pointer;

begin
  WriteBuffer(Value, SizeOf(OleVariant));
  if VarIsArray(Value) then  {cf 1.13.13}
  begin
    {see section 4.6.4.6 of DA 2.04 spec}
    {Check that dim count = 1 &&& for now just ignore}
    {make copy of VarArray so we can safely null out the data ptr}
    Move(VD.VArray^, VarArray, SafeArraySize);
    VarArray.Data:= nil;
    WriteBuffer(VarArray, SafeArraySize);
    {write strings}
    if (VD.VType and varTypeMask) = VT_BSTR then
    begin
      for i:= VarArrayLowBound(Value, 1) to VarArrayHighBound(Value, 1) do
        WriteString(Value[i])
    end else
    begin
      Ptr:= VarArrayLock(Value);
      try
        with VD.VArray^ do
          WriteBuffer(Ptr^, ElementSize*Bounds[0].ElementCount)
      finally
        VarArrayUnlock(Value)
      end
    end
  end else
  if VarType(Value) = VT_BSTR then
  begin
    WriteString(Value)
  end
end;

function TDataChangeStream.ItemHeader2(i: Integer): POPCITEMHEADER2;
begin
  Result:= POPCITEMHEADER2(Integer(Memory) + SizeOf(OPCGROUPHEADER) +
               i*ItemHeaderSize)
end;

{ TDataChangeStream1 }

function TDataChangeStream1.ItemHeaderSize: Integer;
begin
  Result:= SizeOf(OPCITEMHEADER1)
end;

procedure TDataChangeStream1.SetTimestamp(i: Integer;
  const aTimestamp: TFiletime);
begin
  ItemHeader1(i)^.ftTimeStampItem:= aTimestamp
end;

function TDataChangeStream1.Format: TDa1Format;
begin
  Result:= da1DataTime
end;

function TDataChangeStream1.ItemHeader1(i: Integer): POPCITEMHEADER1;
begin
  Result:= POPCITEMHEADER1(ItemHeader2(i))
end;

{ TDataChangeStream2 }

function TDataChangeStream2.Format: TDa1Format;
begin
  Result:= da1Data
end;

function TDataChangeStream2.ItemHeaderSize: Integer;
begin
  Result:= SizeOf(OPCITEMHEADER2)
end;

constructor EOpcError.Create(aErrorCode: HRESULT);
begin
  inherited Create('');
  FErrorCode:= aErrorCode
end;


function EOpcError.GetErrorCode: HRESULT;
begin
  if Succeeded(FErrorCode) then
    Result:= E_FAIL
  else
    Result:= FErrorCode
end;

{ TEnumerateCP }

function TEnumerateCP.Clone(out Enum: IEnumConnectionPoints): HResult;
begin
  Enum:= TEnumerateCP.Create(CP, FIndex);
  Result:= S_OK
end;

constructor TEnumerateCP.Create(aCP: IConnectionPoint; aIndex: Integer);
begin
  inherited Create;
  CP:= aCP;
  FIndex:= aIndex
end;

function TEnumerateCP.Next(celt: Integer; out elt;
  pceltFetched: PLongint): HResult;
var
  Items: Integer;
begin
  Items:= 1 - FIndex;
  Pointer(elt):= nil;
  if (Items = 1) and (celt > 0) then
  begin
    IUnknown(elt):= CP;
    if Assigned(pceltFetched) then
      pceltFetched^:= 1;
    Inc(FIndex)
  end;
  if celt = Items then
    Result:= S_OK
  else
    Result:= S_FALSE
end;

function TEnumerateCP.Reset: HResult;
begin
  FIndex:= 0;
  Result:= S_OK
end;

function TEnumerateCP.Skip(celt: Integer): HResult;
begin
  if (FIndex = 0) and (celt > 0) then
  begin
    Inc(FIndex);
    if celt = 1 then
      Result:= S_OK
    else
      Result:= S_FALSE
  end else
  begin
    if celt = 0 then
      Result:= S_OK
    else
      Result:= S_FALSE
  end
end;

{ TItemProperty }


{ TNamespaceNode }

function TNamespaceNode.Child(i: Integer): TNamespaceNode;
begin
  Result:= nil
end;

function TNamespaceNode.ChildCount: Integer;
begin
  Result:= 0
end;

constructor TNamespaceNode.Create(aParent: TItemIdList;
  const aName: String);
begin
  inherited Create;
  FParent:= aParent;
  FName:= aName;
  if Assigned(aParent) then
    aParent.AddChild(Self)
end;

function TNamespaceNode.Path: string;
var
  delim: Char;

  procedure AddParent(aParent: TNamespaceNode);
  begin
    if Assigned(aParent) and (aParent.Name <> '') then
    begin
      Result:= aParent.Name + delim + Result;
      AddParent(aParent.Parent)
    end
  end;

begin
  delim:= GOpcItemServer.PathDelimiter;
  Result:= Name;
  AddParent(FParent)
end;

{ TItemIdList }

function TItemIdList.AddItemID(const ItemID: String;
  AccessRights: TAccessRights; VarType: Integer): TNamespaceNode;
{note: This will raise EListError if ItemId is already in list. This
will be an unknown exception in OPC client}

{in a hierarchical space a full path can be added}
var
  iid: String;
  Delimiter: Char;
  P, P1: PChar;
  S: string;
  i: Integer;
  Obj: TObject;
  Branch: TItemIdList absolute Obj;
begin
  if GOpcItemServer.HierarchicalBrowsing then
  begin
    Delimiter:= GOpcItemServer.PathDelimiter;
    Branch:= Self;
    P:= PChar(ItemID);
    P1:= P;
    while P^ <> #0 do
    begin
      if P^ = Delimiter then
      begin
        SetString(S, P1, P - P1);
        Branch:= Branch.NewBranch(S);
        P1:= P + 1
      end;
      Inc(P)
    end;
    SetString(S, P1, P - P1);
    iid:= S
  end else
  begin
    iid:= ItemId;
    Branch:= Self
  end;
  if Branch.FChildren.Find(S, i) then   {cf 1.14.30}
  begin
    if not GOpcItemServer.FIgnoreDuplicatesInListItemIds then
      raise EListError.CreateResFmt(@SDuplicateItemId, [S]);
    Result:= TNamespaceNode(Branch.FChildren.Objects[i])
  end else
  begin
    Result:= TNamespaceItem.Create(Branch, iid, AccessRights, VarType)
  end
end;

function TItemIdList.Child(i: Integer): TNamespaceNode;
begin
  Result:= TNamespaceNode(FChildren.Objects[i])
end;

function TItemIdList.ChildCount: Integer;
begin
  Result:= FChildren.Count
end;

procedure TItemIdList.Clear;
var
  i: Integer;
begin
  if Assigned(FChildren) then
  begin
    for i:= 0 to FChildren.Count - 1 do
      FChildren.Objects[i].Free;
    FChildren.Clear
  end
end;

constructor TItemIdList.Create(aParent: TItemIdList;
  const aName: String);
begin
  inherited Create(aParent, aName);
  FChildren:= TStringList.Create;
  with FChildren do
  begin
    Sorted:= true;
    Duplicates:= dupIgnore  {cf 1.15.6.1}
    {Duplicates:= dupError prior to cf 1.15.6.1}
  end
end;

destructor TItemIdList.Destroy;
begin
  Clear;
  FChildren.Free;
  inherited Destroy
end;

function TItemIdList.NewBranch(const aName: String): TItemIDList;
var
  i: Integer;
  Obj: TObject absolute Result;
begin
  if not GOpcItemServer.HierarchicalBrowsing then
    raise EOpcServer.CreateRes(@SCannotCallNewBranchOnFlatSpace);
  if aName = '' then
   raise EOpcServer.CreateResFmt(@SInvalidBranchName, [aName]);
  if FChildren.Find(aName, i) then
  begin
    Obj:= FChildren.Objects[i];
    if not (Obj is TItemIdList) then
      raise EOpcServer.CreateResFmt(@SInvalidBranchName, [aName])
  end else
  begin
    Result:= TItemIdList.Create(Self, aName)
  end
end;

function TItemIdList.Find(const Path: string): TNamespaceNode;
var
  Delimiter: Char;
  P, P1: PChar;
  S: string;
  i: Integer;
  Obj: TObject;
  Branch: TItemIdList absolute Obj;
begin
  try
    if GOpcItemServer.HierarchicalBrowsing then
    begin
      Delimiter:= GOpcItemServer.PathDelimiter;
      Branch:= Self;
      P:= PChar(Path);
      P1:= P;
      while P^ <> #0 do
      begin
        if P^ = Delimiter then
        begin
          SetString(S, P1, P - P1);
          if not Branch.FChildren.Find(S, i) then
            Abort;
          Obj:= Branch.FChildren.Objects[i];
          if not (Obj is TItemIdList) then
            Abort;
          P1:= P + 1
        end;
        Inc(P)
      end;
      SetString(S, P1, P - P1);
      if not Branch.FChildren.Find(S, i) then
        Abort;
      Result:= TNamespaceNode(Branch.FChildren.Objects[i])
    end else
    begin
      if not FChildren.Find(Path, i) then
        Abort;
      Result:= TNamespaceNode(FChildren.Objects[i])
    end
  except
    on EAbort do
      Result:= nil
  end
end;

procedure TItemIDList.AddChild(Child: TNamespaceNode);
{note: Name is duplicated as Entry in FChildren and
as a field of Child. This allows items to be found
efficiently by name (binary search in FChildren) and is
only a bit wasteful of storage as Delphi stores a pointer
only to duplicated strings}
begin
  FChildren.AddObject(Child.Name, Child)
end;

{ TNamespaceItem }

constructor TNamespaceItem.Create(aParent: TItemIdList;
  const aName: String; aAccessRights: TAccessRights; aVarType: Integer);
begin
  inherited Create(aParent, aName);
  FAccessRights:= aAccessRights;
  FVarType:= aVarType
end;

initialization
finalization
  GOpcItemServer.Free;
  GOpcItemServer:= nil
end.

