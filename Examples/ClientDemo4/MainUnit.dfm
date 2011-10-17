object MainForm: TMainForm
  Left = 851
  Top = 384
  Width = 277
  Height = 153
  Caption = 'Very Simple Client'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 16
  object ReadResult: TLabel
    Left = 112
    Top = 88
    Width = 75
    Height = 16
    Caption = 'Read Result'
  end
  object StartButton: TButton
    Left = 16
    Top = 16
    Width = 81
    Height = 25
    Caption = 'Start'
    TabOrder = 0
    OnClick = StartButtonClick
  end
  object WriteButton: TButton
    Left = 16
    Top = 48
    Width = 81
    Height = 25
    Caption = 'Write'
    TabOrder = 1
    OnClick = WriteButtonClick
  end
  object ReadButton: TButton
    Left = 16
    Top = 80
    Width = 81
    Height = 25
    Caption = 'Read'
    TabOrder = 2
    OnClick = ReadButtonClick
  end
  object WriteDataEdit: TEdit
    Left = 112
    Top = 48
    Width = 57
    Height = 24
    TabOrder = 3
    Text = '400'
  end
  object WriteData: TUpDown
    Left = 169
    Top = 48
    Width = 19
    Height = 24
    Associate = WriteDataEdit
    Min = 200
    Max = 800
    Increment = 50
    Position = 400
    TabOrder = 4
    Wrap = False
  end
end
