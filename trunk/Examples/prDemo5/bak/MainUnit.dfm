object MainForm: TMainForm
  Left = 858
  Top = 330
  Width = 536
  Height = 394
  Caption = 'MainForm'
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
    Top = 167
    Width = 528
    Height = 6
    Cursor = crVSplit
    Align = alBottom
    Beveled = True
  end
  object TreeView: TTreeView
    Left = 0
    Top = 41
    Width = 528
    Height = 126
    Align = alClient
    Indent = 19
    ReadOnly = True
    TabOrder = 0
    OnDblClick = TreeViewDblClick
  end
  object DebugLog: TListBox
    Left = 0
    Top = 173
    Width = 528
    Height = 194
    Align = alBottom
    ItemHeight = 16
    TabOrder = 1
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 528
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    object ShutdownClientsButton: TButton
      Left = 16
      Top = 8
      Width = 121
      Height = 25
      Caption = 'Shutdown Clients'
      TabOrder = 0
      OnClick = ShutdownClientsButtonClick
    end
    object ExitServerButton: TButton
      Left = 152
      Top = 8
      Width = 121
      Height = 25
      Caption = 'Exit Server'
      TabOrder = 1
      OnClick = ExitServerButtonClick
    end
  end
end
