object MainForm: TMainForm
  Left = 567
  Top = 392
  Width = 564
  Height = 287
  Caption = 'Opc Client Demo 2'
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
  object GroupLabel: TLabel
    Left = 120
    Top = 8
    Width = 37
    Height = 16
    Caption = 'Group'
  end
  object LogLabel: TLabel
    Left = 120
    Top = 136
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
    Left = 120
    Top = 152
    Width = 417
    Height = 89
    ItemHeight = 16
    TabOrder = 2
  end
  object ItemList: TListView
    Left = 120
    Top = 24
    Width = 417
    Height = 105
    Columns = <
      item
        Caption = 'Item'
        Width = 100
      end
      item
        Caption = 'Value'
        Width = 100
      end
      item
        Caption = 'Quality'
        Width = 100
      end
      item
        Caption = 'Timestamp'
        Width = 100
      end>
    HideSelection = False
    ReadOnly = True
    RowSelect = True
    TabOrder = 3
    ViewStyle = vsReport
    OnChange = ItemListChange
  end
  object BrowseButton: TButton
    Left = 8
    Top = 88
    Width = 81
    Height = 25
    Caption = 'Browse'
    TabOrder = 4
    OnClick = BrowseButtonClick
  end
  object WriteButton: TButton
    Left = 8
    Top = 120
    Width = 81
    Height = 25
    Caption = 'Write'
    TabOrder = 5
    OnClick = WriteButtonClick
  end
  object Client: TOpcSimpleClient
    Groups = <
      item
        Name = 'Group0'
        UpdateRate = 1000
        Active = True
        Items.Strings = (
          'FormHeight'
          'FormWidth'
          'TickCount'
          'TimeOfDay')
        OnDataChange = ClientGroups0DataChange
      end>
    ProgID = 'prDemo5.TDemo5.1'
    ConnectIOPCShutdown = True
    OnConnect = ClientConnect
    OnDisconnect = ClientDisconnect
    OnServerShutdown = ClientServerShutdown
    Left = 16
    Top = 208
  end
end
