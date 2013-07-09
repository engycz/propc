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
unit OpcServerUnit;

interface

uses Windows, SysUtils, Classes, Forms, Variants,
     ActiveX, ComServ, Contnrs,
     prOpcDa, prOpcServer, prOpcTypes, prOpcClasses, prOpcError, prOpcItems;

const
  ServerGuid: TGUID = '{210B0FF5-5E3A-4EE7-B4C3-1494FADD5349}';
  ServerVersion = 1;
  ServerDesc = 'prOpcKit - Compliance Test Demo';
  ServerVendor = 'Production Robots Eng. Ltd.';

type
  TDemoComplianceTest = class(TOpcDataItemServer)
  private
    function OnWrite(Sender : TOPCDataItem; Value : OleVariant):Boolean;

  public
    OPC_Timer : TOPCDataItem;
    OPCItemList : TObjectList;
    constructor Create;
    destructor Destroy; override;
    procedure ListItemIds(List: TItemIDList); override;
  end;

implementation

{ TDemoComplianceTest }

procedure TDemoComplianceTest.ListItemIds(List: TItemIDList);
 procedure ListItemIds_(List: TItemIdList);
 var
   I : Integer;
 begin
   for I := 0 to OPCItemList.Count-1 do
    AddOPCItem(List, TOPCDataItem(OPCItemList[i]).Descr, TOPCDataItem(OPCItemList[i]));
 end;
begin
  ListItemIds_(List);
end;

constructor TDemoComplianceTest.Create;
 procedure AddItem(VarType : Integer; Descr : string);
 var
   I : TOPCDataItem;
 begin
   I := TOPCDataItem.Create(VarType, True, 'Unit', Descr);
   I.OnWrite := OnWrite;
{   if VarType in [varDouble, varSingle] then
    I.EUInfo := TAnalogLimits.Create(-5,100);}
   OPCItemList.Add(I);
   OPCItemList.Add(TOPCDataItem.Create(VarType, True, 'Unit', Descr+'RO'));
 end;
begin
  inherited;
  OPCItemList := TObjectList.Create;
  AddItem(varSmallint, 'SmallInt');//= $0002; { vt_i2           2 }
  AddItem(varInteger,  'Integer');//= $0003; { vt_i4           3 }
  AddItem(varSingle,   'Single');//= $0004; { vt_r4           4 }
  AddItem(varDouble,   'Double');//= $0005; { vt_r8           5 }
  AddItem(varCurrency, 'Currency');//= $0006; { vt_cy           6 }
  AddItem(varDate,     'Date');//= $0007; { vt_date         7 }
  AddItem(varOleStr,   'OleStr');//= $0008; { vt_bstr         8 }
  AddItem(varBoolean,  'Boolean');//= $000B; { vt_bool        11 }
  AddItem(varShortInt, 'ShortInt');//= $0010; { vt_i1          16 }
  AddItem(varByte,     'Byte');//= $0011; { vt_ui1         17 }
  AddItem(varWord,     'Word');//= $0012; { vt_ui2         18 }
  AddItem(varLongWord, 'LongWord');//= $0013; { vt_ui4         19 }*)

  AddItem(varSmallint, 'Branch.SmallInt');//= $0002; { vt_i2           2 }
  AddItem(varInteger,  'Branch.Integer');//= $0003; { vt_i4           3 }
  AddItem(varSingle,   'Branch.Single');//= $0004; { vt_r4           4 }
  AddItem(varDouble,   'Branch.Double');//= $0005; { vt_r8           5 }
  AddItem(varCurrency, 'Branch.Currency');//= $0006; { vt_cy           6 }
  AddItem(varDate,     'Branch.Date');//= $0007; { vt_date         7 }
  AddItem(varOleStr,   'Branch.OleStr');//= $0008; { vt_bstr         8 }
  AddItem(varBoolean,  'Branch.Boolean');//= $000B; { vt_bool        11 }
  AddItem(varShortInt, 'Branch.ShortInt');//= $0010; { vt_i1          16 }
  AddItem(varByte,     'Branch.Byte');//= $0011; { vt_ui1         17 }
  AddItem(varWord,     'Branch.Word');//= $0012; { vt_ui2         18 }
  AddItem(varLongWord, 'Branch.LongWord');//= $0013; { vt_ui4         19 }*)
end;

destructor TDemoComplianceTest.Destroy;
begin
  OPCItemList.Free;
  inherited;
end;

function TDemoComplianceTest.OnWrite(Sender: TOPCDataItem;
  Value: OleVariant): Boolean;
begin
  Sender.Value := Value;
  Result := True;
end;

initialization
  RegisterOPCServer(ServerGUID, ServerVersion, ServerDesc, ServerVendor, TDemoComplianceTest.Create)
end.

