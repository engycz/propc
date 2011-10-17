object MainForm: TMainForm
  Left = 509
  Top = 339
  Width = 536
  Height = 201
  Caption = 'Demo13 - Enumerated EU with Rtti'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 16
  object DebugLog: TListBox
    Left = 0
    Top = 41
    Width = 528
    Height = 128
    Align = alClient
    ItemHeight = 16
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 528
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
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
