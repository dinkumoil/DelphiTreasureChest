object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'Single App Instance'
  ClientHeight = 464
  ClientWidth = 586
  Color = clBtnFace
  DefaultMonitor = dmPrimary
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  DesignSize = (
    586
    464)
  PixelsPerInch = 96
  TextHeight = 13
  object btnClose: TButton
    Left = 488
    Top = 431
    Width = 90
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Schlie'#223'en'
    TabOrder = 0
    OnClick = btnCloseClick
  end
  object memOutput: TMemo
    Left = 8
    Top = 8
    Width = 570
    Height = 409
    Anchors = [akLeft, akTop, akRight, akBottom]
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
  end
end
