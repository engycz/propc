object MainForm: TMainForm
  Left = 596
  Top = 356
  Width = 457
  Height = 314
  Caption = 'Deadband and Enumerated Demo'
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
  object LogLabel: TLabel
    Left = 16
    Top = 152
    Width = 64
    Height = 16
    Caption = 'Messages'
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
    Left = 16
    Top = 176
    Width = 409
    Height = 89
    ItemHeight = 16
    TabOrder = 2
  end
  object TrackBar: TTrackBar
    Left = 104
    Top = 8
    Width = 305
    Height = 41
    Orientation = trHorizontal
    Frequency = 5
    Position = 0
    SelEnd = 0
    SelStart = 0
    TabOrder = 3
    TickMarks = tmBottomRight
    TickStyle = tsAuto
    OnChange = TrackBarChange
  end
  object ListBox: TListBox
    Left = 144
    Top = 64
    Width = 225
    Height = 97
    ItemHeight = 16
    TabOrder = 4
    OnClick = ListBoxClick
  end
  object Client: TOpcSimpleClient
    Groups = <
      item
        Name = 'Group0'
        UpdateRate = 1000
        Active = True
        PercentDeadband = 10
        Items.Strings = (
          'TrackBar'
          'ListBox')
        OnDataChange = ClientGroups0DataChange
      end>
    ProgID = 'prDemo15.TDemo15.1'
    SortedItemLists = False
    ConnectIOPCShutdown = True
    OnConnect = ClientConnect
    OnDisconnect = ClientDisconnect
    OnServerShutdown = ClientServerShutdown
    Left = 16
    Top = 96
  end
end
