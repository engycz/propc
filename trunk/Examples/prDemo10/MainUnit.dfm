object MainForm: TMainForm
  Left = 798
  Top = 376
  Width = 312
  Height = 258
  Caption = 'Demo10 - Arrays '
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 304
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
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
  object UpdateButton: TButton
    Left = 112
    Top = 192
    Width = 75
    Height = 25
    Caption = 'Update'
    TabOrder = 1
    OnClick = UpdateButtonClick
  end
  object Edit1: TEdit
    Left = 8
    Top = 64
    Width = 81
    Height = 24
    TabOrder = 2
  end
  object Edit2: TEdit
    Left = 8
    Top = 96
    Width = 81
    Height = 24
    TabOrder = 3
  end
  object Edit3: TEdit
    Left = 8
    Top = 128
    Width = 81
    Height = 24
    TabOrder = 4
  end
  object Edit4: TEdit
    Left = 8
    Top = 160
    Width = 81
    Height = 24
    TabOrder = 5
  end
  object Edit5: TEdit
    Left = 8
    Top = 192
    Width = 81
    Height = 24
    TabOrder = 6
  end
end
