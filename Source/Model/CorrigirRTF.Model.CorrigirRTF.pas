unit CorrigirRTF.Model.CorrigirRTF;

interface

uses
  WPCTRRich,
  WPRTEDefsConsts,
  WPRTEDefs,
  Vcl.Controls,
  Vcl.StdCtrls,
  Vcl.Forms,
  Vcl.Graphics,
  Data.DB,
  System.SysUtils,
  System.Classes,
  System.Generics.Collections,
  SimpleDS,
  Vcl.ComCtrls,
  Vcl.Dialogs,
  FireDAC.Comp.Client;

type
  IRegistros = interface
    ['{BE0180A7-CAEE-4B31-9783-93EBF60EA813}']
    function CorririgrRegistros(pDataSet: TFDQuery; pDataSetCampos: TFDQuery; pProgressBar: TProgressBar):IRegistros;
    procedure ExibirMensagemConfirmacao;
  end;

  ICorrigirRTF = interface
    ['{2FE254F6-DD30-45CB-AEEE-2D44FAB72601}']
    function ListarTodosBookMarks:ICorrigirRTF;
    function RemoverBookMarks: ICorrigirRTF;
    function IncluirBooksMarks: ICorrigirRTF;
    procedure PersistirBaseDeDados(pDataSet: TFDQuery; pNomeCampoRTF: string);
  end;

  TDadosBookMark = class
    Chave             : string;
    Nome              : string;
    Display           : string;
    PosicaoOriginal   : Integer;
    PosicaoTemporaria : Integer;
    Tamanho           : Integer;
  end;

  TCorrigirRTF= class(TInterfacedObject,ICorrigirRTF, IRegistros)
  private
    FWptEditor        : TWPRichText;
    FDadosBookMark    : TDadosBookMark;
    FListaBookMarks   : TWPTextObjList;
    FListaDeBookMarks : TStringList;
    procedure ConfigurarWpTools;
    procedure AdicionaListaBookMark(pLista: TStringList; pDados: TDadosBookMark);
  public
    constructor Create;
    class function New: IRegistros;
    destructor Destroy;override;
    function IncluirBooksMarks: ICorrigirRTF;
    function ListarTodosBookMarks: ICorrigirRTF;
    function RemoverBookMarks: ICorrigirRTF;
    function CorririgrRegistros(pDataSet: TFDQuery; pDataSetCampos: TFDQuery; pProgressBar: TProgressBar): IRegistros;
    procedure PersistirBaseDeDados(pDataSet: TFDQuery; pNomeCampoRTF: string);
    procedure ExibirMensagemConfirmacao;
  end;

implementation

{ TCorrigirRTF }

procedure TCorrigirRTF.AdicionaListaBookMark(pLista: TStringList; pDados: TDadosBookMark);
Var
 RecNum : Integer;
 Guid   : TGuid;
 sChave : string;
begin
  CreateGUID(Guid);
  sChave := GUIDToString(Guid);

  pLista.AddObject(sChave,TDadosBookMark.Create);
  RecNum := pred(pLista.Count);

  with TDadosBookMark(pLista.Objects[RecNum]) do
  begin
    Chave             := sChave;
    Nome              := pDados.Nome;
    Display           := pDados.Display;
    PosicaoOriginal   := pDados.PosicaoOriginal;
    PosicaoTemporaria := pDados.PosicaoTemporaria;
    Tamanho           := pDados.Tamanho;
  end;
end;

procedure TCorrigirRTF.ConfigurarWpTools;
begin
  FWptEditor.Cursor                              := crIBeam;
  FWptEditor.Header.DefaultTabstop               := 254;
  FWptEditor.PrintParameter.PrintOptions         := [wpDoNotChangePrinterDefaults];
  FWptEditor.XOffset                             := 10;
  FWptEditor.YOffset                             := 141;
  FWptEditor.ScrollBars                          := ssVertical;
  FWptEditor.EditOptions                         := [wpTableResizing, wpTableColumnResizing, wpObjectMoving, wpObjectResizingWidth, wpObjectResizingHeight, wpObjectResizingKeepRatio,
                                                     wpObjectSelecting, wpObjectDeletion, wpSpreadsheetCursorMovement, wpActivateUndo, wpActivateUndoHotkey, wpMoveCPOnPageUpDown];
  FWptEditor.ProtectedProp                       := [ppParProtected, ppProtected, ppProtectSelectedTextToo];
  FWptEditor.BorderStyle                         := bsSingle;
  FWptEditor.HyperLinkCursor                     := crDefault;
  FWptEditor.TextObjectCursor                    := crArrow;
  FWptEditor.InsertPointAttr.TextColor           := clRed;
  FWptEditor.InsertPointAttr.UseTextColor        := True;
  FWptEditor.HyperlinkTextAttr.Underline         := tsFALSE;
  FWptEditor.HyperlinkTextAttr.HotStyleIsActive  := True;
  FWptEditor.HyperlinkTextAttr.UnderlineColor    := clBlue;
  FWptEditor.HyperlinkTextAttr.TextColor         := clOlive;
  FWptEditor.HyperlinkTextAttr.UseUnderlineColor := True;
  FWptEditor.HyperlinkTextAttr.UseTextColor      := True;
  FWptEditor.HyperlinkTextAttr.HotUnderlineColor := clRed;
  FWptEditor.HyperlinkTextAttr.HotTextColor      := clRed;
  FWptEditor.HyperlinkTextAttr.HotUnderline      := tsTRUE;
  FWptEditor.BookmarkTextAttr.UnderlineColor     := clGreen;
  FWptEditor.BookmarkTextAttr.TextColor          := clRed;
  FWptEditor.BookmarkTextAttr.UseUnderlineColor  := True;
  FWptEditor.BookmarkTextAttr.UseTextColor       := True;
  FWptEditor.HiddenTextAttr.Hidden               := True;
  FWptEditor.ProtectedTextAttr.TextColor         := clBlue;
  FWptEditor.ProtectedTextAttr.UseTextColor      := True;
  FWptEditor.ViewOptions                         := [wpShowPageNRinGap, wpDrawFineUnderlines];
  FWptEditor.ParentColor                         := False;
  FWptEditor.TabStop                             := True;
end;

function TCorrigirRTF.CorririgrRegistros(pDataSet: TFDQuery; pDataSetCampos: TFDQuery; pProgressBar: TProgressBar): IRegistros;
begin
  Result := Self;

  try
    pProgressBar.Min := 0;
    pProgressBar.Max := pDataSet.RecordCount;

    pDataSet.First;
    while not pDataSet.Eof do
    begin
      pDataSetCampos.Close;
      pDataSetCampos.SQL.Clear;
      pDataSetCampos.SQL.Add(Format('select %s from %s where %s is not null',[pDataSet.FieldByName('Campo').AsString,
                                                                              pDataSet.FieldByName('Tabela').AsString,
                                                                              pDataSet.FieldByName('Campo').AsString]));
      pDataSetCampos.Open;

      pDataSetCampos.First;
      while not pDataSetCampos.eof do
      begin
        FWptEditor.AsString := pDataSetCampos.FieldByName(pDataSet.FieldByName('Campo').AsString).AsString;

        Self.ListarTodosBookMarks
            .RemoverBookMarks
            .IncluirBooksMarks
            .PersistirBaseDeDados(pDataSetCampos, pDataSet.FieldByName('Campo').AsString);

        pDataSetCampos.Next;
      end;
      pDataSet.Next;
      pProgressBar.Position := pDataSet.RecNo;
    end;
    pDataSet.First;

    pProgressBar.Position := 0;
  except
    on E:Exception do
    begin
      raise Exception.Create(E.Message);
    end;
  end;
end;

constructor TCorrigirRTF.Create;
begin
  inherited Create;
  FListaDeBookMarks := TStringList.Create;
  FDadosBookMark    := TDadosBookMark.Create;
  FWptEditor        := TWPRichText.Create(nil);

  ConfigurarWpTools;
end;

destructor TCorrigirRTF.Destroy;
begin
  FreeAndNil(FDadosBookMark);
  FreeAndNil(FWptEditor);
  FreeAndNil(FListaDeBookMarks);
  inherited;
end;

procedure TCorrigirRTF.ExibirMensagemConfirmacao;
begin
  ShowMessage('Minutas Atualizadas com Sucesso');
end;

function TCorrigirRTF.IncluirBooksMarks: ICorrigirRTF;
var
  iContador    : Integer;
  PosicaoFinal : Integer;
begin
  Result := Self;

  FWptEditor.CPPosition := 0;
  for iContador := 0 to pred(FListaDeBookMarks.Count) do
  begin
    FWptEditor.CPPosition := TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).PosicaoOriginal;

    if TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Display = EmptyStr then
    begin
      while FWptEditor.CPChar <> #$D do
      begin
        FWptEditor.CPMoveNext;
      end;

      FWptEditor.CPMoveBack;
      PosicaoFinal := FWptEditor.CPPosition;
      FWptEditor.SetSelPosLen(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).PosicaoOriginal,PosicaoFinal);
      FWptEditor.BookmarkInput(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Nome,False);
    end
    else
    begin
      FWptEditor.InputString(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Display);
      FWptEditor.SetSelPosLen(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).PosicaoOriginal,TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Tamanho);
      FWptEditor.BookmarkInput(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Nome,False);
    end;
    FWptEditor.HideSelection;
  end;
end;

function TCorrigirRTF.ListarTodosBookMarks: ICorrigirRTF;
var
  iContador           : Integer;
  iAcrescentarPosicao : Integer;
  iPosicao            : Integer;
begin
  Result := Self;
  iAcrescentarPosicao := 0;
  FListaBookMarks := TWPTextObjList.Create;
  try
    FWptEditor.BookmarkGetList(FListaBookMarks);

    FListaDeBookMarks.Clear;

    for iContador := 0 to Pred(FListaBookMarks.Count) do
    begin
      FWptEditor.BookmarkMoveToNext(FListaBookMarks[iContador].Name,False);
      iPosicao                         := FWptEditor.CPPosition;
      FDadosBookMark.Nome              := FListaBookMarks[iContador].Name;
      FDadosBookMark.Display           := FListaBookMarks[iContador].EmbeddedText;
      FDadosBookMark.PosicaoOriginal   := iposicao;
      FDadosBookMark.Tamanho           := FListaBookMarks[iContador].EmbeddedTextLength;
      FDadosBookMark.PosicaoTemporaria := iAcrescentarPosicao + (iPosicao -FDadosBookMark.Tamanho)  ;

      iAcrescentarPosicao := iAcrescentarPosicao + (Length(FListaBookMarks[iContador].Name) -1);
      AdicionaListaBookMark(FListaDeBookMarks,FDadosBookMark);
    end;
  finally
    FreeAndNil(FListaBookMarks);
  end;
end;

class function TCorrigirRTF.New: IRegistros;
begin
  Result := Self.Create;
end;

procedure TCorrigirRTF.PersistirBaseDeDados(pDataSet: TFDQuery; pNomeCampoRTF: string);
begin
  pDataSet.Edit;
  pDataSet.FieldByName(pNomeCampoRTF).AsString := FWptEditor.AsString;
  pDataSet.Post;

//  pDataSet.ApplyUpdates(-1);
end;

function TCorrigirRTF.RemoverBookMarks: ICorrigirRTF;
var
  iContador: Integer;
begin
  Result := Self;

  FWptEditor.CPPosition := 0;
  for iContador := 0 to Pred(FListaDeBookMarks.Count) do
  begin
    FWptEditor.SetSelPosLen(FWptEditor.BookmarkFind(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Nome),TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Tamanho +1);
    FWptEditor.ClearSelection;
    FWptEditor.BookmarkDelete(TDadosBookMark(FListaDeBookMarks.Objects[FListaDeBookMarks.IndexOf(FListaDeBookMarks[iContador])]).Nome, True,True);
  end;
end;

end.

