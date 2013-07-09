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
