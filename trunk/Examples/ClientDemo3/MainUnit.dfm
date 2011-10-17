object MainForm: TMainForm
  Left = 20
  Top = 523
  Width = 203
  Height = 87
  Caption = 'Very Simple Client'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -10
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object StartButton: TButton
    Left = 13
    Top = 7
    Width = 67
    Height = 20
    Caption = 'Start'
    TabOrder = 0
    OnClick = StartButtonClick
  end
end
