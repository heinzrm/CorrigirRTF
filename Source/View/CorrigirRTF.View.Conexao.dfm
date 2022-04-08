object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Corrigir BookMarks Das minutas'
  ClientHeight = 157
  ClientWidth = 678
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 678
    Height = 140
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object gbProtesto: TGroupBox
      Left = 0
      Top = 0
      Width = 678
      Height = 105
      Align = alTop
      Caption = 'Conex'#227'o Protesto:'
      TabOrder = 0
      object lblNomeBancoDeDados: TLabel
        Left = 252
        Top = 43
        Width = 82
        Height = 13
        Caption = 'Banco De Dados:'
      end
      object lblUsuario: TLabel
        Left = 20
        Top = 67
        Width = 40
        Height = 13
        Caption = 'Usu'#225'rio:'
      end
      object lblServidor: TLabel
        Left = 16
        Top = 43
        Width = 44
        Height = 13
        Caption = 'Servidor:'
      end
      object lblSenha: TLabel
        Left = 300
        Top = 67
        Width = 34
        Height = 13
        Caption = 'Senha:'
      end
      object lblTipoBanco: TLabel
        Left = 4
        Top = 19
        Width = 56
        Height = 13
        Caption = 'Tipo Banco:'
      end
      object edtUsuarioProtesto: TEdit
        Left = 66
        Top = 66
        Width = 168
        Height = 19
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 0
      end
      object edtServidorProtesto: TEdit
        Left = 66
        Top = 42
        Width = 169
        Height = 19
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 1
      end
      object edtSenhaProtesto: TEdit
        Left = 338
        Top = 66
        Width = 151
        Height = 19
        Ctl3D = False
        ParentCtl3D = False
        PasswordChar = '*'
        TabOrder = 2
      end
      object edtBancoProtesto: TEdit
        Left = 338
        Top = 42
        Width = 297
        Height = 19
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 3
      end
      object cbTipoBanco: TComboBox
        Left = 66
        Top = 16
        Width = 145
        Height = 21
        TabOrder = 4
        Items.Strings = (
          'Sql Server'
          'Firebird')
      end
    end
    object Panel2: TPanel
      Left = 0
      Top = 110
      Width = 678
      Height = 30
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 1
      DesignSize = (
        678
        30)
      object btnExecutar: TButton
        Left = 599
        Top = 1
        Width = 75
        Height = 25
        Anchors = [akTop, akRight]
        Caption = 'Executar'
        TabOrder = 0
        OnClick = btnExecutarClick
      end
    end
  end
  object ProgressBar1: TProgressBar
    Left = 0
    Top = 140
    Width = 678
    Height = 17
    Align = alBottom
    TabOrder = 1
  end
  object sdsMinutas: TSimpleDataSet
    Aggregates = <>
    DataSet.CommandText = 'select * from minutas'
    DataSet.MaxBlobSize = -1
    DataSet.Params = <>
    Params = <>
    Left = 152
    Top = 80
  end
  object sdsTabelaIni: TSimpleDataSet
    Aggregates = <>
    DataSet.CommandText = 'select observacao from ini where observacao is not null'
    DataSet.MaxBlobSize = -1
    DataSet.Params = <>
    Params = <>
    Left = 232
    Top = 82
  end
  object FDTabelas: TFDQuery
    Connection = FDConexao
    Left = 416
    Top = 96
  end
  object FDConexao: TFDConnection
    Left = 528
    Top = 96
  end
  object FDCampoRTF: TFDQuery
    Connection = FDConexao
    Left = 344
    Top = 96
  end
end
