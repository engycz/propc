{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
{------------------------------------------------------------}
unit prOpcEnum;
{$I prOpcCompilerDirectives.inc}
interface
uses
  SysUtils, Classes, Windows, prOpcDa, ComObj, ActiveX, prOpcTypes,
  prOpcComn;

type
  TServerSearchOption = (soIncludeOldServers, soUseOpcEnum);
  TServerSearchOptions = set of TServerSearchOption;

  TOpcServerInfo = record
    Clsid: TGUID;
    ProgID: String;
    UserType: String;
    VerIndProgID: String;
    Vendor: String;
    DaTypes: TOpcDataAccessTypes;
  end;

  TOpcServerList = array of TOpcServerInfo;

  EOpcEnum = class(Exception);


procedure GetOpcServers(const Host: String; var Servers: TOpcServerList;
  Options: TServerSearchOptions = []);


procedure GetServerInfo(const Host, ProgID: String;
                        var Clsid: TGUID;
                        var UserType: String;
                        var Vendor: String;
                        Options: TServerSearchOptions = []);

function IsLocalHost(const Host: String): Boolean;

implementation
uses
  Registry;


resourcestring
  SCouldNotConnect = 'Could not connect to registry on %s';
  SOnRemote = ' on %s';
  SCannotFindServer = 'Cannot find server %s';
  SClassNotRegistered = 'Class not registered';


function IsLocalHost(const Host: String): Boolean;
var
  Buf: array [0..MAX_COMPUTERNAME_LENGTH] of Char;
  BufSize: Cardinal;
begin
  Result:= true;
  if Host <> '' then
  begin
    BufSize:= SizeOf(Buf);
    if not GetComputerName(Buf, BufSize) then
{$IFDEF D6UP}
      RaiseLastOSError
{$ELSE}
      RaiseLastWin32Error
{$ENDIF}
    else
      Result:= SameText(Host, Buf)
  end
end;

function OpenReg(const Host: String): TRegistry;
var
  UncName: String;
begin
  Result:= TRegistry.Create(KEY_READ);  
  Result.RootKey:= HKEY_LOCAL_MACHINE;
  if not IsLocalHost(Host) then
  begin
    if Copy(Host, 1, 2) <> '\\' then
      UncName:= '\\' + Host
    else
      UncName:= Host;
    if not Result.RegistryConnect(UncName) then
    begin
       Result.Free;  {cf 1.16b.1}
      raise ERegistryException.CreateFmt(SCouldNotConnect, [Host]);
    end;
  end;
end;

const
  BaseKey = '\Software\Classes';
  AllocBy = 20;

procedure GetOpcServers(const Host: String; var Servers: TOpcServerList;
  Options: TServerSearchOptions);

type
  PServerListRec = ^TOpcServerInfo;

var
  AllocCount: Integer;
  ServerCount: Integer;

function IncServerCount: PServerListRec;
begin
  Inc(ServerCount);
  if ServerCount > AllocCount then
  begin
    Inc(AllocCount, AllocBy);
    SetLength(Servers, AllocCount)
  end;
  Result:= @Servers[ServerCount-1]
end;


function IndexOfClsid(const Clsid: TGUID): Integer;
var
  i: Integer;
begin
  Result:= -1;
  for i:= 0 to ServerCount - 1 do
  if IsEqualGUID(Clsid, Servers[i].Clsid) then
  begin
    Result:= i;
    break
  end
end;



{these searches could be optimised by rewriting using the API -
We iterate through the registry twice, where we could do it once}
procedure SearchForServers; {without using OpcEnum}
var
  R: TRegistry;
  CatID: array[TOpcDataAccessType] of String;
  KeyNames: TStrings;

{new servers use component categories to identify
themselves}
procedure CheckForNewServer(const Clsid: String);
var
  ClsGuid: TGUID;
  P: PServerListRec;
  PushKey: String;
  DaTypes: TOpcDataAccessTypes;
  Da: TOpcDataAccessType;
begin
  DaTypes:= [];
  for Da:= Low(TOpcDataAccessType) to High(TOpcDataAccessType) do
  if R.KeyExists(Clsid + '\Implemented Categories\' + CatID[Da]) then
    Include(DaTypes, Da);
  if DaTypes <> [] then
  begin
    ClsGuid:= StringToGUID(Clsid);
    P:= IncServerCount;
    P^.Clsid:= ClsGuid;
    P^.DaTypes:= DaTypes;
    PushKey:= '\' + R.CurrentPath;
    try
      R.OpenKeyReadOnly(Clsid);
      P^.UserType:= R.ReadString('');
      R.OpenKeyReadOnly('ProgID');
      P^.ProgID:= R.ReadString('');

      if R.OpenKeyReadOnly(PushKey + '\' + Clsid + '\VersionIndependentProgID') then
        P^.VerIndProgID := R.ReadString('')
      else
        P^.VerIndProgID := P^.ProgID;

      if R.OpenKeyReadOnly(BaseKey + '\' + P^.ProgID + '\Opc\Vendor') then
        P^.Vendor:= R.ReadString('')
    finally
      R.OpenKeyReadOnly(PushKey)
    end
  end
end;

{old servers don't use component categories, so we will have to search for
the 'opc' key. This search will be very slow}
procedure CheckForOldServer(const ProgID: String);
var
  ClsGuid: TGUID;
  P: PServerListRec;
  PushKey: String;
  UserType: String;
  i: Integer;
begin
  if (Length(ProgID) > 0) and (ProgID[1] <> '.') then {ignore file extensions}
  begin
    if R.KeyExists(ProgID + '\Opc') and
       R.KeyExists(ProgID + '\Clsid') then  {make sure this is a com object}
    begin
      PushKey:= '\' + R.CurrentPath;
      try
        R.OpenKeyReadOnly(ProgID);
        UserType:= R.ReadString('');
        R.OpenKeyReadOnly('Clsid');
        ClsGuid:= StringToGuid(R.ReadString(''));
        i:= IndexOfClsid(ClsGuid);
        if i = -1 then
        begin
          New(P);
          P^.Clsid:= ClsGuid;
          P^.ProgID:= ProgID;
          P^.UserType:= UserType;
          P^.DaTypes:= [OPCDA1];  {this is a bit of a guess}
          if R.OpenKeyReadOnly(PushKey + '\' + ProgID + '\Opc\Vendor') then
            P^.Vendor:= R.ReadString('');
        end
      finally
        R.OpenKeyReadOnly(PushKey)
      end
    end
  end
end;

var
  i: Integer;

begin
  CatID[opcDA1]:= GuidToString(CATID_OPCDAServer10);
  CatID[opcDA2]:= GuidToString(CATID_OPCDAServer20);
  CatID[opcDA3]:= GuidToString(CATID_OPCDAServer30);
  KeyNames:= TStringList.Create;
  R:= OpenReg(Host);
  try
    {look for implemented categories}
    R.OpenKeyReadOnly(BaseKey + '\Clsid');
    R.GetKeyNames(KeyNames);
    for i:= 0 to KeyNames.Count - 1 do
      CheckForNewServer(KeyNames[i]);
    if soIncludeOldServers in Options then  {this will be VERY slow on remote machines}
    begin
      R.OpenKeyReadOnly(BaseKey);
      KeyNames.Clear;
      R.GetKeyNames(KeyNames); {there may be squillions}
      for i:= 0 to KeyNames.Count - 1 do
        CheckForOldServer(KeyNames[i])
    end;
  finally
    R.Free;
    KeyNames.Free
  end
end;

procedure EnumServers;
var
//  List: IOPCServerList;
//  Enum: IEnumGUID;
  List: IOPCServerList2;
  Enum: IOPCEnumGUID;
  NextClsid: TGUID;

const
  CatIDs: array[TOpcDataAccessType] of PGUID =
    (@CATID_OPCDAServer10, @CATID_OPCDAServer20, @CATID_OPCDAServer30);

procedure AddClasses(DaType: TOpcDataAccessType);
var
  Fetched: DWORD;
  Res: PServerListRec;
  szUserType: PWideChar;
  szProgID: PWideChar;
  szVerIndProgID: PWideChar;
  i: Integer;
begin
  OleCheck(List.EnumClassesOfCategories(1, PGUIDList(CatIDs[DaType]), 0, nil, Enum));
  while Enum.Next(1, NextClsid, Fetched) = S_OK do
  begin
    i:= IndexOfClsid(NextClsid);
    if i <> -1 then
    begin
      Res:= @Servers[i];
      Include(Res^.DaTypes, DaType)
    end else
    begin
      try
        OleCheck(List.GetClassDetails(NextClsid, szProgID, szUserType, szVerIndProgID));
        Res:= IncServerCount;
        with Res^ do
        begin
          Clsid:= NextClsid;
          ProgID:= WideCharToString(szProgID);
          UserType:= WideCharToString(szUserType);
          VerIndProgID:= WideCharToString(szVerIndProgID);
          Vendor:= '';
          DaTypes:= [DaType]
        end;
        CoTaskMemFree(szUserType);
        CoTaskMemFree(szProgID);
        CoTaskMemFree(szVerIndProgID);
      except
      end;
    end
  end
end;

begin
  if IsLocalHost(Host) then
    List:= CreateComObject(CLSID_OPCServerList) as IOPCServerList2
  else
    List:= CreateRemoteComObject(Host, CLSID_OPCServerList) as IOPCServerList2;
  AddClasses(opcDA1);
  AddClasses(opcDA2);
  AddClasses(opcDA3)
end;

begin
  ServerCount:= 0;
  AllocCount:= 0;
  if soUseOpcEnum in Options then
    EnumServers
  else
    SearchForServers;
  SetLength(Servers, ServerCount)
end;

procedure GetServerInfoWithOpcEnum(const Host, ProgID: String;
                               var Clsid: TGUID;
                               var UserType: String);
var
//  List: IOPCServerList;
  List: IOPCServerList2;
  WProgID: WideString;
  PUserType: PWideChar;
  PProgID: PWideChar;
  PVerIndProgID: PWideChar;
begin
  if IsLocalHost(Host) then
    List:= CreateComObject(CLSID_OPCServerList) as IOPCServerList2
  else
    List:= CreateRemoteComObject(Host, CLSID_OPCServerList) as IOPCServerList2;
  WProgID:= ProgID;
  List.CLSIDFromProgID(PWideChar(WProgID), Clsid);
  List.GetClassDetails(Clsid, PProgID, PUserType, PVerIndProgID);
  UserType:= PUserType;
  CoTaskMemFree(PUserType);
  CoTaskMemFree(PProgID);
  CoTaskMemFree(PVerIndProgID);
end;

procedure GetServerInfo(const Host, ProgID: String;
                        var Clsid: TGUID;
                        var UserType: String;
                        var Vendor: String;
                        Options: TServerSearchOptions = []);

procedure Bomb(const Msg: String);
var
  ErrMsg: String;
begin
  FmtStr(ErrMsg, Msg, [ProgID]);
  if Host <> '' then
    ErrMsg:= ErrMsg + Format(SOnRemote, [Host]);
  raise EOpcEnum.Create(ErrMsg)
end;

var
  R: TRegistry;
  PushKey: String;
begin
  if soUseOpcEnum in Options then
  begin
    Vendor:= '';
    GetServerInfoWithOpcEnum(Host, ProgID, Clsid, UserType)
  end else
  begin
    R:= OpenReg(Host);
    try
      if not R.OpenKeyReadOnly(BaseKey + '\' + ProgID) then
        Bomb(SCannotFindServer);
      UserType:= R.ReadString('');
      PushKey:= '\' + R.CurrentPath;
      if not R.OpenKeyReadOnly('Clsid') then
        Bomb(SClassNotRegistered);
      Clsid:= StringToGuid(R.ReadString(''));
      R.OpenKeyReadOnly(PushKey);
      if R.OpenKeyReadOnly('Opc\Vendor') then
        Vendor:= R.ReadString('')
      else
        Vendor:= ''
    finally
      R.Free
    end
  end
end;

end.


