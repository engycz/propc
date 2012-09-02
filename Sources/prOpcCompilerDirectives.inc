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

{$DEFINE D6UP}

// Test Delphi 5
{$IFDEF VER130}
  {$UNDEF D6UP}
{$ENDIF}

{$IFDEF CONDITIONALEXPRESSIONS}
  // Enable Unicode strings in RTTI
  {$IF CompilerVersion >= 22.0}
    {$DEFINE F_UString}
  {$IFEND}
{$ENDIF}

