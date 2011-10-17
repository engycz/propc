object MainForm: TMainForm
  Left = 858
  Top = 330
  Width = 437
  Height = 174
  Caption = 'Analog and Enumerated EU Demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 16
  object ExitServerButton: TButton
    Left = 152
    Top = 16
    Width = 121
    Height = 25
    Caption = 'Exit Server'
    TabOrder = 0
    OnClick = ExitServerButtonClick
  end
  object ShutdownClientsButton: TButton
    Left = 16
    Top = 16
    Width = 121
    Height = 25
    Caption = 'Shutdown Clients'
    TabOrder = 1
    OnClick = ShutdownClientsButtonClick
  end
  object TrackBar: TTrackBar
    Left = 160
    Top = 56
    Width = 241
    Height = 45
    Max = 100
    Orientation = trHorizontal
    Frequency = 5
    Position = 0
    SelEnd = 0
    SelStart = 0
    TabOrder = 2
    TickMarks = tmBottomRight
    TickStyle = tsAuto
  end
  object ListBox: TComboBox
    Left = 24
    Top = 56
    Width = 121
    Height = 24
    Style = csDropDownList
    ItemHeight = 16
    ItemIndex = 0
    TabOrder = 3
    Text = 'Red'
    Items.Strings = (
      'Red'
      'Yellow'
      'Blue'
      'Orange'
      'White'
      'Turquoise'
      'Brown')
  end
end
