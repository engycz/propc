object MainForm: TMainForm
  Left = 596
  Top = 356
  Width = 491
  Height = 204
  Caption = 'In-Process Demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object Label1: TLabel
    Left = 105
    Top = 8
    Width = 72
    Height = 17
    Alignment = taRightJustify
    Caption = 'TickCount'
  end
  object Label2: TLabel
    Left = 284
    Top = 8
    Width = 69
    Height = 16
    Alignment = taRightJustify
    Caption = 'TimeOfDay'
  end
  object ConnectButton: TButton
    Left = 8
    Top = 24
    Width = 81
    Height = 25
    Caption = 'Connect'
    TabOrder = 0
    OnClick = ConnectButtonClick
  end
  object DisconnectButton: TButton
    Left = 8
    Top = 56
    Width = 81
    Height = 25
    Caption = 'Disconnect'
    TabOrder = 1
    OnClick = DisconnectButtonClick
  end
  object MsgLog: TListBox
    Left = 104
    Top = 40
    Width = 353
    Height = 97
    ItemHeight = 16
    TabOrder = 2
  end
  object TickCountValue: TStaticText
    Left = 184
    Top = 8
    Width = 89
    Height = 20
    AutoSize = False
    BorderStyle = sbsSingle
    TabOrder = 3
  end
  object TimeOfDayValue: TStaticText
    Left = 360
    Top = 8
    Width = 97
    Height = 20
    AutoSize = False
    BorderStyle = sbsSingle
    TabOrder = 4
  end
  object Client: TOpcSimpleClient
    Groups = <
      item
        Name = 'Group0'
        UpdateRate = 1000
        Active = True
        Items.Strings = (
          'TimeOfDay'
          'TickCount')
        OnDataChange = ClientGroups0DataChange
      end>
    ProgID = 'prdemo6.TDemo6.1'
    SortedItemLists = False
    ConnectIOPCShutdown = False
    OnConnect = ClientConnect
    OnDisconnect = ClientDisconnect
    Left = 16
    Top = 96
  end
end
