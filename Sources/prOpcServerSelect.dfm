object ServerSelectDlg: TServerSelectDlg
  Left = 317
  Top = 258
  BorderStyle = bsDialog
  Caption = 'Select Server'
  ClientHeight = 262
  ClientWidth = 607
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  PixelsPerInch = 120
  TextHeight = 16
  object List: TListView
    Left = 0
    Top = 0
    Width = 607
    Height = 193
    Align = alTop
    Columns = <
      item
        Caption = 'Prog ID'
        Width = 150
      end
      item
        Caption = 'User'
        Width = 150
      end
      item
        Caption = 'DA Support'
        Width = 100
      end
      item
        Caption = 'Vendor'
        Width = 200
      end>
    RowSelect = True
    TabOrder = 0
    ViewStyle = vsReport
  end
  object OKBtn: TButton
    Left = 380
    Top = 214
    Width = 93
    Height = 31
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 1
  end
  object CancelBtn: TButton
    Left = 495
    Top = 214
    Width = 92
    Height = 31
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 2
  end
end
