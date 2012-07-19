{------------------------------------------------------------}
{                    prOpc Toolkit                           }
{ Copyright (c) 2000, 2001 Production Robots Engineering Ltd }
{                                                            }
{ mailto engycz@gmail.com                                    }
{ http://code.google.com/p/propc/                            }
{ original source mailto: prOpcKit@prel.co.uk                }
{ original source http://www.prel.co.uk                      }
{------------------------------------------------------------}
{History

1.14 01/04/02  New File
}

{$IFDEF Evaluation}
  {$L-}
  {$D-}
  {$Y-}
  {$DEFINE NoMasks}
{$ENDIF}

{$IFDEF VER140}
  {$DEFINE D6UP}
{$ENDIF}

{$IFDEF VER150}
  {$DEFINE D6UP}
  {$DEFINE D7UP}
{$ENDIF}

{$IFDEF VER170}
  {$DEFINE D6UP}
  {$DEFINE D7UP}
  {$DEFINE D9UP}
{$ENDIF}

{$IFDEF VER180}
  {$DEFINE D6UP}
  {$DEFINE D7UP}
  {$DEFINE D9UP}
  {$DEFINE D10UP}
{$ENDIF}

//this Delphi 2007 for Win32
{$IFDEF VER185}
  {$DEFINE D6UP}
  {$DEFINE D7UP}
  {$DEFINE D9UP}
  {$DEFINE D10UP}
  {$DEFINE D11UP}
{$ENDIF}

//CodeGear RAD Studio 2007
{$IFDEF VER190}
  {$DEFINE D6UP}
  {$DEFINE D7UP}
  {$DEFINE D9UP}
  {$DEFINE D10UP}
  {$DEFINE D11UP}
{$ENDIF}
