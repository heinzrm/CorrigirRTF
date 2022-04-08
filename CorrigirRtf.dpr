program CorrigirRTF;

uses
  Vcl.Forms,
  CorrigirRTF.View.Conexao in 'Source\View\CorrigirRTF.View.Conexao.pas' {Form1},
  CorrigirRTF.Model.CorrigirRTF in 'Source\Model\CorrigirRTF.Model.CorrigirRTF.pas',
  CorrigirRTF.Conexao.Conexao in 'Source\Conexao\CorrigirRTF.Conexao.Conexao.pas',
  CorrigirRTF.Model.TipoConexao in 'Source\Model\CorrigirRTF.Model.TipoConexao.pas',
  CorrigirRTF.DIversos.Enumerado in 'Source\Diversos\CorrigirRTF.DIversos.Enumerado.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

