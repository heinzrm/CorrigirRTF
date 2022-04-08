unit CorrigirRTF.Conexao.Conexao;

interface

uses
  System.IniFiles,
  System.SysUtils,
  Vcl.Dialogs,
  FireDAC.Comp.Client;{, Cripto;}

type
  TConexao = class
    private
    FServidor: string;
    FDriver: string;
    FPath: string;
    FAuthSistema: string;
    FSenha: string;
    FDatabase: string;
    FNomeConexao: string;
    FUsuario: string;
    FPorta: Integer;
    FSecao: string;
//    FCripto: TCriptografa;
    procedure SetAuthSistema(const Value: string);
    procedure SetDatabase(const Value: string);
    procedure SetDriver(const Value: string);
    procedure SetNomeConexao(const Value: string);
    procedure SetPath(const Value: string);
    procedure SetPorta(const Value: Integer);
    procedure SetSecao(const Value: string);
    procedure SetSenha(const Value: string);
    procedure SetServidor(const Value: string);
    procedure SetUsuario(const Value: string);
    public
      property Path:string read FPath write SetPath;
      property Servidor:string read FServidor write SetServidor;
      property Porta:Integer read FPorta write SetPorta;
      property Database:string read FDatabase write SetDatabase;
      property Senha:string read FSenha write SetSenha;
      property Usuario:string read FUsuario write SetUsuario;
      property Driver:string read FDriver write SetDriver;
      property Secao:string read FSecao write SetSecao;
      property NomeConexao:string read FNomeConexao write SetNomeConexao;
      property AuthSistema:string read FAuthSistema write SetAuthSistema;

      constructor Create(Path:string;Secao:string);
      destructor Destroy;override;
      procedure LeIni();virtual;
      procedure GravaIni(Usuario,Senha,Servidor,Banco:string;Porta:Integer);virtual;
      procedure Conectar(var Conexao:TFDConnection);virtual;
  end;
implementation

{ TConexao }

procedure TConexao.Conectar(var Conexao: TFDConnection);
begin
  LeIni();
  try
    Conexao.Connected := False;
    Conexao.LoginPrompt := False;
    Conexao.Params.Clear;
    Conexao.Params.Add('Server='+ Servidor);
    Conexao.Params.Add('user_name='+ Usuario);
    Conexao.Params.Add('password='+ Senha);
    Conexao.Params.Add('port='+ IntToStr(Porta));
    Conexao.Params.Add('Database='+ Database);
    Conexao.Params.Add('DriverID='+ Driver);
    Conexao.Params.Add('ConnectionName='+NomeConexao);
    Conexao.Params.Add('OSAuthent='+AuthSistema);
  Except
    on E:Exception do
      ShowMessage('Erro ao carregar parâmetros de conexão!'#13#10 + E.Message);
  end;
end;

constructor TConexao.Create(Path, Secao: string);
begin
//  FCripto := TCriptografa.Create(nil);
//  FCripto.Key := 'PescaMilagrosa';

  if FileExists(Path) then
  begin
    Self.Path := Path;
    Self.Secao := Secao;
  end
  else
     raise Exception.Create('Arquivo INI para configuração não encontrado.'#13#10'Aplicação será finalizada.');
end;

destructor TConexao.Destroy;
begin
//  FreeAndNil(FCripto);
  inherited;
end;

procedure TConexao.GravaIni(Usuario, Senha, Servidor, Banco: string;
  Porta: Integer);
var
  ArqIni : TIniFile;
begin
  ArqIni := TIniFile.Create(Path);
  try
    ArqIni.WriteString(Secao, 'Usuario', Usuario);
    ArqIni.WriteString(Secao, 'Senha', Senha);
    ArqIni.WriteString(Secao, 'Database', Banco);
    ArqIni.WriteString(Secao, 'Servidor', Servidor);
    ArqIni.WriteInteger(Secao, 'Porta', Porta);
    ArqIni.WriteString(Secao, 'DriverName', Driver);
  finally
     ArqIni.Free;
  end;
end;

procedure TConexao.LeIni;
var
  ArqIni : TIniFile;
begin
  ArqIni := TIniFile.Create(Path);
  try
    Servidor    := ArqIni.ReadString(Secao, 'Servidor', '');
    Porta       := ArqIni.ReadInteger(Secao, 'Porta', 0);
    Database    := ArqIni.ReadString(Secao, 'Database', '');
    Senha       := ArqIni.ReadString(Secao, 'Password', '');
    Usuario     := ArqIni.ReadString(Secao, 'User_Name', '');
    Driver      := ArqIni.ReadString(Secao, 'DriverName', '');
    NomeConexao := ArqIni.ReadString(Secao, 'NomeConexao', '');
    AuthSistema := ArqIni.ReadString(Secao, 'OSAuthent', '');

//    Servidor    := FCripto.CriptoHexToText(ArqIni.ReadString(Secao, 'Servidor', ''));
//    Porta       := ArqIni.ReadInteger(Secao, 'Porta', 0);
//    Database    := FCripto.CriptoHexToText(ArqIni.ReadString(Secao, 'Database', ''));
//    Senha       := FCripto.CriptoHexToText(ArqIni.ReadString(Secao, 'Password', ''));
//    Usuario     := FCripto.CriptoHexToText(ArqIni.ReadString(Secao, 'User_Name', ''));
//    Driver      := ArqIni.ReadString(Secao, 'DriverName', '');
//    NomeConexao := ArqIni.ReadString(Secao, 'NomeConexao', '');
//    AuthSistema := ArqIni.ReadString(Secao, 'OSAuthent', '');
  finally
     ArqIni.Free;
  end;
end;

procedure TConexao.SetAuthSistema(const Value: string);
begin
  FAuthSistema := Value;
end;

procedure TConexao.SetDatabase(const Value: string);
begin
  FDatabase := Value;
end;

procedure TConexao.SetDriver(const Value: string);
begin
  FDriver := Value;
end;

procedure TConexao.SetNomeConexao(const Value: string);
begin
  FNomeConexao := Value;
end;

procedure TConexao.SetPath(const Value: string);
begin
  FPath := Value;
end;

procedure TConexao.SetPorta(const Value: Integer);
begin
  FPorta := Value;
end;

procedure TConexao.SetSecao(const Value: string);
begin
  FSecao := Value;
end;

procedure TConexao.SetSenha(const Value: string);
begin
  FSenha := Value;
end;

procedure TConexao.SetServidor(const Value: string);
begin
  FServidor := Value;
end;

procedure TConexao.SetUsuario(const Value: string);
begin
  FUsuario := Value;
end;

end.

