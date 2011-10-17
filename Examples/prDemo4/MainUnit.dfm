object MainForm: TMainForm
  Left = 529
  Top = 356
  Width = 536
  Height = 394
  Caption = 'Demo4 - Client hook demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 16
  object Splitter1: TSplitter
    Left = 0
    Top = 205
    Width = 528
    Height = 6
    Cursor = crVSplit
    Align = alBottom
    Beveled = True
  end
  object TreeView: TTreeView
    Left = 0
    Top = 0
    Width = 528
    Height = 205
    Align = alClient
    Indent = 19
    ReadOnly = True
    TabOrder = 0
    OnDblClick = TreeViewDblClick
  end
  object DebugLog: TListBox
    Left = 0
    Top = 211
    Width = 528
    Height = 151
    Align = alBottom
    ItemHeight = 16
    TabOrder = 1
  end
end
