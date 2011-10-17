object MainForm: TMainForm
  Left = 703
  Top = 377
  Width = 307
  Height = 94
  Caption = 'Client for prDemo13'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 16
  object ConnectButton: TButton
    Left = 16
    Top = 16
    Width = 97
    Height = 25
    Caption = 'Connect'
    TabOrder = 0
    OnClick = ConnectButtonClick
  end
  object BorderStyleCB: TComboBox
    Left = 136
    Top = 16
    Width = 145
    Height = 24
    Style = csDropDownList
    Enabled = False
    ItemHeight = 16
    TabOrder = 1
    OnSelect = BorderStyleCBSelect
  end
  object Client: TOpcSimpleClient
    Groups = <
      item
        Name = 'Group'
        UpdateRate = 0
        Active = False
        Items.Strings = (
          'BorderStyle')
      end>
    ProgID = 'prDemo13.TDemo13.1'
    SortedItemLists = False
    ConnectIOPCShutdown = False
    OnConnect = ClientConnect
    OnDisconnect = ClientDisconnect
  end
end
