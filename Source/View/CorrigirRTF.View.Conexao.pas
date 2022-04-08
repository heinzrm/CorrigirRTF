unit CorrigirRTF.View.Conexao;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Data.DB,
  Datasnap.DBClient,
  SimpleDS,
  Vcl.ExtCtrls,
  Vcl.StdCtrls,
  WPRTEDefs,
  WPCTRMemo,
  WPCTRRich,
  Data.SqlExpr,
  WPTbar,
  System.Threading,
  Vcl.ComCtrls,
  CorrigirRTF.Conexao.Conexao, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, FireDAC.UI.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Phys, FireDAC.VCLUI.Wait,
  FireDAC.Comp.Client, FireDAC.Comp.DataSet, FireDAC.Phys.MSSQLDef,
  FireDAC.Phys.ODBCBase, FireDAC.Phys.MSSQL, CorrigirRTF.Model.TipoConexao,
  CorrigirRTF.DIversos.Enumerado;

type
  TDadosBookMark = class
    Chave             : string;
    Nome              : string;
    Display           : string;
    PosicaoOriginal   : Integer;
    PosicaoTemporaria : Integer;
    Tamanho           : Integer;
  end;


  TForm1 = class(TForm)
    btnExecutar: TButton;
    Panel1: TPanel;
    sdsMinutas: TSimpleDataSet;
    ProgressBar1: TProgressBar;
    gbProtesto: TGroupBox;
    lblNomeBancoDeDados: TLabel;
    lblUsuario: TLabel;
    lblServidor: TLabel;
    lblSenha: TLabel;
    edtUsuarioProtesto: TEdit;
    edtServidorProtesto: TEdit;
    edtSenhaProtesto: TEdit;
    edtBancoProtesto: TEdit;
    Panel2: TPanel;
    sdsTabelaIni: TSimpleDataSet;
    cbTipoBanco: TComboBox;
    lblTipoBanco: TLabel;
    FDTabelas: TFDQuery;
    FDConexao: TFDConnection;
    FDCampoRTF: TFDQuery;
    procedure btnExecutarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    ClassConexao: TConexao;
    procedure CarregarConsulta(pQuery: TFDQuery);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  CorrigirRTF.Model.CorrigirRTF;

{$R *.dfm}

procedure TForm1.btnExecutarClick(Sender: TObject);
begin
  CarregarConsulta(FDTabelas);

  TCorrigirRTF.New
              .CorririgrRegistros(FDTabelas,FDCampoRTF,ProgressBar1)
              .ExibirMensagemConfirmacao;

  sdsMinutas.Close;
end;

procedure TForm1.CarregarConsulta(pQuery: TFDQuery);
begin
  case TTipoSGDB(cbTipoBanco.ItemIndex) of
    tcSqlServer:
    begin
      pQuery.SQL.Clear;
      pQuery.SQL.Add('SELECT distinct');
      pQuery.SQL.Add('[Tabela] = OBJECT_NAME(c.object_id),');
      pQuery.SQL.Add('[Campo] = c.name');
      pQuery.SQL.Add('FROM sys.columns c');
      pQuery.SQL.Add('INNER JOIN sys.systypes tp on tp.xtype = c.system_type_id');
      pQuery.SQL.Add('LEFT OUTER JOIN sys.extended_properties ex ON ex.major_id = c.object_id');
      pQuery.SQL.Add('AND ex.minor_id = c.column_id');
      pQuery.SQL.Add('AND ex.name = ''MS_Description''');
      pQuery.SQL.Add('where tp.name = ''Image''');
      pQuery.SQL.Add('ORDER BY OBJECT_NAME(c.object_id)');
    end;
    tcFirebird:
    begin
      pQuery.SQL.Clear;
      pQuery.SQL.Add('SELECT');
      pQuery.SQL.Add('RDB$RELATION_FIELDS.RDB$RELATION_NAME TABLE_NAME Tabela,');
      pQuery.SQL.Add('RDB$RELATION_FIELDS.RDB$FIELD_NAME FIELD_NAME Campo');
      pQuery.SQL.Add('FROM');
      pQuery.SQL.Add('RDB$RELATION_FIELDS');
      pQuery.SQL.Add('JOIN RDB$FIELDS');
      pQuery.SQL.Add('ON RDB$FIELDS.RDB$FIELD_NAME = RDB$RELATION_FIELDS.RDB$FIELD_SOURCE');
      pQuery.SQL.Add('JOIN RDB$TYPES');
      pQuery.SQL.Add('ON RDB$FIELDS.RDB$FIELD_TYPE = RDB$TYPES.RDB$TYPE AND');
      pQuery.SQL.Add('RDB$TYPES.RDB$FIELD_NAME = ''RDB$FIELD_TYPE''');
      pQuery.SQL.Add('where  RDB$TYPES.RDB$TYPE_NAME = ''BLOB''');
      pQuery.SQL.Add('ORDER BY');
      pQuery.SQL.Add('RDB$RELATION_FIELDS.RDB$RELATION_NAME');
    end;
  end;
  pQuery.Open;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  ClassConexao := TConexao.Create(Format('%s\Configuracao.ini',[ExtractFileDir(ParamStr(0))]),'GERAL');
  ClassConexao.Conectar(FDConexao);

  edtServidorProtesto.Text := FDConexao.Params.Values['Server'];
  edtBancoProtesto.Text    := FDConexao.Params.Values['Database'];
  edtUsuarioProtesto.Text  := FDConexao.Params.Values['user_name'];
  edtSenhaProtesto.Text    := FDConexao.Params.Values['password'];

  cbTipoBanco.ItemIndex := Integer(TTipoConexao.New.RetornaTipoBanco(FDConexao.Params.Values['DriverID']));
end;

end.


