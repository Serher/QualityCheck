object Form1: TForm1
  Left = 717
  Top = 186
  Width = 440
  Height = 510
  Caption = 'Quality Check'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 193
    Height = 19
    Caption = #1040#1076#1088#1077#1089#1089' '#1086#1090#1095#1077#1090#1072' '#1080#1079' '#1073#1080#1083#1083#1080#1085#1075#1072
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clGreen
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 16
    Top = 72
    Width = 150
    Height = 19
    Caption = #1040#1076#1088#1077#1089' '#1086#1073#1097#1077#1075#1086' '#1086#1090#1095#1077#1090#1072
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clGreen
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label3: TLabel
    Left = 16
    Top = 144
    Width = 109
    Height = 19
    Caption = #1061#1086#1076' '#1074#1099#1087#1086#1083#1085#1077#1085#1080#1103
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = [fsItalic]
    ParentFont = False
  end
  object Button1: TButton
    Left = 320
    Top = 32
    Width = 89
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 16
    Top = 32
    Width = 297
    Height = 27
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clBlue
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
  end
  object Memo1: TMemo
    Left = 16
    Top = 168
    Width = 393
    Height = 241
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object Button2: TButton
    Left = 16
    Top = 416
    Width = 393
    Height = 41
    Caption = #1053#1040#1063#1040#1058#1068' '#1055#1056#1054#1042#1045#1056#1050#1059
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clRed
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = [fsBold, fsItalic]
    ParentFont = False
    TabOrder = 3
    OnClick = Button2Click
  end
  object Edit2: TEdit
    Left = 16
    Top = 96
    Width = 297
    Height = 27
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clBlue
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
  end
  object Button3: TButton
    Left = 320
    Top = 96
    Width = 89
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 288
    Top = 152
    Width = 123
    Height = 17
    Caption = #1055#1086#1089#1095#1080#1090#1072#1090#1100' '#1089#1090#1072#1090#1080#1089#1090#1080#1082#1091
    TabOrder = 6
    OnClick = Button4Click
  end
  object OpenDialog1: TOpenDialog
    Left = 208
    Top = 136
  end
end
