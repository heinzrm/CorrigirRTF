unit CorrigirRTF.Model.TipoConexao;

interface

uses
  CorrigirRTF.DIversos.Enumerado,
  System.StrUtils, System.SysUtils;

type
  ITipoConexao = interface
    ['{2CAF3010-3D60-49B5-AC6C-B3BB0ECF1F09}']
    function RetornaTipoBanco(pDriverName: string): TTipoSGDB;
  end;

  TTipoConexao = class(TInterfacedObject,ITipoConexao)
  public
    class function New: ITipoConexao;
    function RetornaTipoBanco(pDriverName: string): TTipoSGDB;
  end;
implementation

{ TTipoConexao }

class function TTipoConexao.New: ITipoConexao;
begin
  Result := Self.Create;
end;

function TTipoConexao.RetornaTipoBanco(pDriverName: string): TTipoSGDB;
begin
  Result := tcDesconhecido;

  case AnsiIndexStr(AnsiUppercase(pDriverName),['MSSQL','Firebird']) of
    0: result := TTipoSGDB.tcSqlServer;
    1: result := TTipoSGDB.tcFirebird;
  end;
end;

end.
