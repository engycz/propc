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
{------------------------------------------------------------}
{ This unit derived substantially from work by:              }
{ OPC Programmers' Connection                                }
{ http://www.opcconnect.com/                                 }
{ mailto:opc@dial.pipex.com                                  }
{------------------------------------------------------------}
unit prOpcError;
{$I prOpcCompilerDirectives.inc}
interface
uses
  Windows, SysUtils, ActiveX;

{
Module Name:
    OpcError.h
Author:
    OPC Task Force
Revision History:
Release 1.0A
     Removed Unused messages
     Added OPC_S_INUSE, OPC_E_INVALIDCONFIGFILE, OPC_E_NOTFOUND
Release 2.0
     Added OPC_E_INVALID_PID
  mgl added string conversion function to delphi translation 27/11/00
}

{
Code Assignements:
  0000 to 0200 are reserved for Microsoft use
  (although some were inadverdantly used for OPC Data Access 1.0 errors).
  0200 to 7FFF are reserved for future OPC use.
  8000 to FFFF can be vendor specific.
}


const

  //
  //  Values are 32 bit values laid out as follows:
  //
  //   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
  //   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
  //  +---+-+-+-----------------------+-------------------------------+
  //  |Sev|C|R|     Facility          |               Code            |
  //  +---+-+-+-----------------------+-------------------------------+
  //
  //  where
  //
  //      Sev - is the severity code
  //
  //          00 - Success
  //          01 - Informational
  //          10 - Warning
  //          11 - Error
  //
  //      C - is the Customer code flag
  //
  //      R - is a reserved bit
  //
  //      Facility - is the facility code
  //
  //      Code - is the facility's status code
  //

  // OPC Data Access

  //
  // MessageId: OPC_E_INVALIDHANDLE
  //
  // MessageText:
  //
  //  The value of the handle is invalid.
  //
  OPC_E_INVALIDHANDLE = HResult($C0040001);

  //
  // MessageId: OPC_E_BADTYPE
  //
  // MessageText:
  //
  //  The server cannot convert the data between the
  //  requested data type and the canonical data type.
  //
  OPC_E_BADTYPE = HResult($C0040004);

  //
  // MessageId: OPC_E_PUBLIC
  //
  // MessageText:
  //
  //  The requested operation cannot be done on a public group.
  //
  OPC_E_PUBLIC = HResult($C0040005);

  //
  // MessageId: OPC_E_BADRIGHTS
  //
  // MessageText:
  //
  //  The Items AccessRights do not allow the operation.
  //
  OPC_E_BADRIGHTS = HResult($C0040006);

  //
  // MessageId: OPC_E_UNKNOWNITEMID
  //
  // MessageText:
  //
  //  The item is no longer available in the server address space.
  //
  OPC_E_UNKNOWNITEMID = HResult($C0040007);

  //
  // MessageId: OPC_E_INVALIDITEMID
  //
  // MessageText:
  //
  //  The item definition doesn't conform to the server's syntax.
  //
  OPC_E_INVALIDITEMID = HResult($C0040008);

  //
  // MessageId: OPC_E_INVALIDFILTER
  //
  // MessageText:
  //
  //  The filter string was not valid.
  //
  OPC_E_INVALIDFILTER = HResult($C0040009);

  //
  // MessageId: OPC_E_UNKNOWNPATH
  //
  // MessageText:
  //
  //  The item's access path is not known to the server.
  //
  OPC_E_UNKNOWNPATH = HResult($C004000A);

  //
  // MessageId: OPC_E_RANGE
  //
  // MessageText:
  //
  //  The value was out of range.
  //
  OPC_E_RANGE = HResult($C004000B);

  //
  // MessageId: OPC_E_DUPLICATENAME
  //
  // MessageText:
  //
  //  Duplicate name not allowed.
  //
  OPC_E_DUPLICATENAME = HResult($C004000C);

  //
  // MessageId: OPC_S_UNSUPPORTEDRATE
  //
  // MessageText:
  //
  //  The server does not support the requested data rate
  //  but will use the closest available rate.
  //
  OPC_S_UNSUPPORTEDRATE = HResult($0004000D);

  //
  // MessageId: OPC_S_CLAMP
  //
  // MessageText:
  //
  //  A value passed to WRITE was accepted but the output was clamped.
  //
  OPC_S_CLAMP = HResult($0004000E);

  //
  // MessageId: OPC_S_INUSE
  //
  // MessageText:
  //
  //  The operation cannot be completed because the
  //  object still has references that exist.
  //
  OPC_S_INUSE = HResult($0004000F);

  //
  // MessageId: OPC_E_INVALIDCONFIGFILE
  //
  // MessageText:
  //
  //  The server's configuration file is an invalid format.
  //
  OPC_E_INVALIDCONFIGFILE = HResult($C0040010);

  //
  // MessageId: OPC_E_NOTFOUND
  //
  // MessageText:
  //
  //  The server could not locate the requested object.
  //
  OPC_E_NOTFOUND = HResult($C0040011);

  //
  // MessageId: OPC_E_INVALID_PID
  //
  // MessageText:
  //
  //  The server does not recognise the passed property ID.
  //
  OPC_E_INVALID_PID = HResult($C0040203);

  //
  // MessageId: OPC_E_DEADBANDNOTSET
  //
  // MessageText:
  //
  //  The item deadband has not been set for this item.
  //
  OPC_E_DEADBANDNOTSET = HResult($C0040400);

  //
  // MessageId: OPC_E_DEADBANDNOTSUPPORTED
  //
  // MessageText:
  //
  //  The item does not support deadband.
  //
  OPC_E_DEADBANDNOTSUPPORTED = HResult($C0040401);

  //
  // MessageId: OPC_E_NOBUFFERING
  //
  // MessageText:
  //
  //  The server does not support buffering of data items that are collected at
  //  a faster rate than the group update rate.
  //
  OPC_E_NOBUFFERING = HResult($C0040402);

  //
  // MessageId: OPC_E_INVALIDCONTINUATIONPOINT
  //
  // MessageText:
  //
  //  The continuation point is not valid.
  //
  OPC_E_INVALIDCONTINUATIONPOINT = HResult($C0040403);

  //
  // MessageId: OPC_S_DATAQUEUEOVERFLOW
  //
  // MessageText:
  //
  //  Data Queue Overflow - Some value transitions were lost.
  //
  OPC_S_DATAQUEUEOVERFLOW = HResult($00040404);

  //
  // MessageId: OPC_E_RATENOTSET
  //
  // MessageText:
  //
  //  Server does not support requested rate.
  //
  OPC_E_RATENOTSET = HResult($C0040405);

  //
  // MessageId: OPC_E_NOTSUPPORTED
  //
  // MessageText:
  //
  //  The server does not support writing of quality and/or timestamp.
  //
  OPC_E_NOTSUPPORTED = HResult($C0040406);


function StdOpcErrorToStr(Code: HRESULT; var Res: String): Boolean;
{return true if found}

implementation

resourcestring
  S_OPC_E_INVALIDHANDLE = 'The value of the handle is invalid.';
  S_OPC_E_BADTYPE = 'The server cannot convert the data between the'+#13+
                    'requested data type and the canonical data type.';
  S_OPC_E_PUBLIC = 'The requested operation cannot be done on a public group.';
  S_OPC_E_BADRIGHTS = 'The Items AccessRights do not allow the operation.';
  S_OPC_E_UNKNOWNITEMID = 'The item is no longer available in the server address space.';
  S_OPC_E_INVALIDITEMID = 'The item definition doesn''t conform to the server''s syntax.';
  S_OPC_E_INVALIDFILTER = 'The filter string was not valid.';
  S_OPC_E_UNKNOWNPATH = 'The item''s access path is not known to the server.';
  S_OPC_E_RANGE = 'The value was out of range.';
  S_OPC_E_DUPLICATENAME = 'Duplicate name not allowed.';
  S_OPC_S_UNSUPPORTEDRATE =  'The server does not support the requested data rate but will use the closest available rate.';
  S_OPC_S_CLAMP = 'A value passed to WRITE was accepted but the output was clamped.';
  S_OPC_S_INUSE = 'The operation cannot be completed because the object still has references that exist.';
  S_OPC_E_INVALIDCONFIGFILE = 'The server''s configuration file is an invalid format.';
  S_OPC_E_NOTFOUND = 'The server could not locate the requested object.';
  S_OPC_E_INVALID_PID = 'The server does not recognise the passed property ID.';
  S_OPC_E_DEADBANDNOTSET = 'The item deadband has not been set for this item.';
  S_OPC_E_DEADBANDNOTSUPPORTED = 'The item does not support deadband.';
  S_OPC_E_NOBUFFERING = 'The server does not support buffering of data items that are collected at a faster rate than the group update rate.';
  S_OPC_E_INVALIDCONTINUATIONPOINT = 'The continuation point is not valid.';
  S_OPC_S_DATAQUEUEOVERFLOW = 'Data Queue Overflow - Some value transitions were lost.';
  S_OPC_E_RATENOTSET = 'Server does not support requested rate.';
  S_OPC_E_NOTSUPPORTED = 'The server does not support writing of quality and/or timestamp.';

function StdOpcErrorToStr(Code: HRESULT; var Res: string): Boolean;
begin
  Result:= true;
  case Code of
    OPC_E_INVALIDHANDLE: Res:= S_OPC_E_INVALIDHANDLE;
    OPC_E_BADTYPE: Res:= S_OPC_E_BADTYPE;
    OPC_E_PUBLIC: Res:= S_OPC_E_PUBLIC;
    OPC_E_BADRIGHTS: Res:= S_OPC_E_BADRIGHTS;
    OPC_E_UNKNOWNITEMID: Res:= S_OPC_E_UNKNOWNITEMID;
    OPC_E_INVALIDITEMID: Res:= S_OPC_E_INVALIDITEMID;
    OPC_E_INVALIDFILTER: Res:= S_OPC_E_INVALIDFILTER;
    OPC_E_UNKNOWNPATH: Res:= S_OPC_E_UNKNOWNPATH;
    OPC_E_RANGE: Res:= S_OPC_E_RANGE;
    OPC_E_DUPLICATENAME: Res:= S_OPC_E_DUPLICATENAME;
    OPC_S_UNSUPPORTEDRATE: Res:= S_OPC_S_UNSUPPORTEDRATE;
    OPC_S_CLAMP: Res:= S_OPC_S_CLAMP;
    OPC_S_INUSE: Res:= S_OPC_S_INUSE;
    OPC_E_INVALIDCONFIGFILE: Res:= S_OPC_E_INVALIDCONFIGFILE;
    OPC_E_NOTFOUND: Res:= S_OPC_E_NOTFOUND;
    OPC_E_INVALID_PID: Res:= S_OPC_E_INVALID_PID;
    OPC_E_DEADBANDNOTSET: Res:= S_OPC_E_DEADBANDNOTSET;
    OPC_E_DEADBANDNOTSUPPORTED: Res:= S_OPC_E_DEADBANDNOTSUPPORTED;
    OPC_E_NOBUFFERING: Res:= S_OPC_E_NOBUFFERING;
    OPC_E_INVALIDCONTINUATIONPOINT: Res:= S_OPC_E_INVALIDCONTINUATIONPOINT;
    OPC_S_DATAQUEUEOVERFLOW: Res:= S_OPC_S_DATAQUEUEOVERFLOW;
    OPC_E_RATENOTSET: Res:= S_OPC_E_RATENOTSET;
    OPC_E_NOTSUPPORTED: Res:= S_OPC_E_NOTSUPPORTED;
  else
    Result:= false
  end
end;


end.
