unit uActiveScript;

interface

uses
  System.SysUtils, System.Classes, System.Math, WinApi.Windows, Winapi.ActiveX,
  Generics.Collections, System.Win.ComObj,

  AscrLib;

type
  TActiveScript = class;
  TScriptErrorEvent = procedure(Sender: TActiveScript; pscripterror: IActiveScriptError) of Object;

  EActiveScriptException = class(EOleException)
  private
    FLineNumber: Cardinal;
    FCharacterPosition: Integer;
  public
    property LineNumber: Cardinal read FLineNumber write FLineNumber;
    property CharacterPosition: Integer read FCharacterPosition write FCharacterPosition;
    constructor Create(const Message: string; ErrorCode: HRESULT;
      const Source, HelpFile: string; HelpContext: Integer; LineNumber: Cardinal;
      CharacterPosition: Integer);
  end;

  // Wrapper component for IActiveScript
  TActiveScript = class(TComponent)
  private
    FOnScriptError: TScriptErrorEvent;
    FExceptionOnError: Boolean;
    FLanguage: string;
    procedure SetLanguage(const Value: string);
    { Private declarations }
  protected
    { Protected declarations }
    type
      TScriptSite = class(TInterfacedObject, IActiveScriptSite,
        IActiveScriptSiteWindow)
      private
        FHwnd: wireHWND;
        FOwner: TActiveScript;
      public
        // IActiveScriptSite
        function GetLCID(out plcid: LongWord): HResult; stdcall;
        function GetItemInfo(pstrName: PWideChar; dwReturnMask: LongWord;
          {out} ppiunkItem, ppti: {PUnknown}PPointer): HResult; stdcall;
        function GetDocVersionString(out pbstrVersion: WideString)
          : HResult; stdcall;
        function OnScriptTerminate(var pvarResult: OleVariant;
          var pexcepinfo: EXCEPINFO): HResult; stdcall;
        function OnStateChange(ssScriptState: tagSCRIPTSTATE): HResult; stdcall;
        function OnScriptError(const pscripterror: IActiveScriptError)
          : HResult; stdcall;
        function OnEnterScript: HResult; stdcall;
        function OnLeaveScript: HResult; stdcall;
        // IActiveScriptSiteWindow
        function GetWindow(out phwnd: wireHWND): HResult; stdcall;
        function EnableModeless(fEnable: Integer): HResult; stdcall;
        constructor Create(AOwner: TActiveScript);
        destructor Destroy(); override;
      end;

  protected
    ScriptSite: TScriptSite;
    Script: IActiveScript;
    ScriptParse: IActiveScriptParse;
    FGlobalObjects: TDictionary<string, {IDispatch}Pointer>;
    procedure CreateEngine();
    procedure ReleaseEngine();
  public const
    DefaultLanguage = 'JScript';
    // Hardcode for now
    CLSID_VBScript: TGUID = '{b54f3741-5b07-11cf-a4b0-00aa004a55e8}';
    CLSID_JScript: TGUID = '{f414c260-6ac0-11cf-b6d1-00aa00bbbb58}';
  public
    { Public declarations }
    HasErrorInfo: Boolean;
    ErrInfo: EXCEPINFO;
    ErrLnNum: Cardinal;
    ErrCharPos: Integer;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy(); override;
    procedure Reset();
    procedure AddObject(const Name: string; const Object_: IDispatch; AddMembers: Boolean);
    function Eval(const Expression: string): OleVariant;
  published
    { Published declarations }
    property ExceptionOnError: Boolean read FExceptionOnError write FExceptionOnError default True;
    property Language: string read FLanguage write SetLanguage;  // Ignored for now
    property OnScriptError: TScriptErrorEvent read FOnScriptError write FOnScriptError;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DWF', [TActiveScript]);
end;

{ TActiveScript.TScriptSite }

constructor TActiveScript.TScriptSite.Create(AOwner: TActiveScript);
begin
  inherited Create();
  FOwner := AOwner;
end;

destructor TActiveScript.TScriptSite.Destroy;
begin

  inherited;
end;

function TActiveScript.TScriptSite.EnableModeless(fEnable: Integer): HResult;
begin
  Result := S_OK;
end;

function TActiveScript.TScriptSite.GetDocVersionString(out pbstrVersion
  : WideString): HResult;
begin
  pbstrVersion := '1.0';
  Result := S_OK;
end;

function TActiveScript.TScriptSite.GetItemInfo(pstrName: PWideChar;
  dwReturnMask: LongWord; {out} ppiunkItem, ppti: {PUnknown}PPointer): HResult;
var
  d: Pointer; //IDispatch;
begin
  if not FOwner.FGlobalObjects.TryGetValue(pstrName, d) then
    Exit(TYPE_E_ELEMENTNOTFOUND);

  if (dwReturnMask and SCRIPTINFO_IUNKNOWN) <> 0 then
  begin
    IDispatch(d)._AddRef();
    ppiunkItem^ := d;
  end;
  if (dwReturnMask and SCRIPTINFO_ITYPEINFO) <> 0 then
  begin
    ppti^ := nil;
  end;
  Result := S_OK;
end;

function TActiveScript.TScriptSite.GetLCID(out plcid: LongWord): HResult;
begin
  plcid := 0;
  Result := S_OK;
end;

function TActiveScript.TScriptSite.GetWindow(out phwnd: wireHWND): HResult;
begin
  phwnd := FHwnd;
  Result := S_OK;
end;

function TActiveScript.TScriptSite.OnEnterScript: HResult;
begin
  Result := S_OK;
end;

function TActiveScript.TScriptSite.OnLeaveScript: HResult;
begin
  Result := S_OK;
end;

function TActiveScript.TScriptSite.OnScriptError(const pscripterror
  : IActiveScriptError): HResult;
var
  Ctx: Cardinal;
begin
  if ((pscripterror.GetExceptionInfo(FOwner.ErrInfo) = S_OK) and
      (pscripterror.GetSourcePosition(Ctx, FOwner.ErrLnNum, FOwner.ErrCharPos) = S_OK)) then
    FOwner.HasErrorInfo := True;
  if Assigned(FOwner.OnScriptError) then
    FOwner.OnScriptError(FOwner, pscripterror);
  Result := S_OK;
end;

function TActiveScript.TScriptSite.OnScriptTerminate(var pvarResult: OleVariant;
  var pexcepinfo: EXCEPINFO): HResult;
begin
  Result := S_OK;
end;

function TActiveScript.TScriptSite.OnStateChange(ssScriptState: tagSCRIPTSTATE)
  : HResult;
begin
  Result := S_OK;
end;


{ TActiveScript }

procedure TActiveScript.AddObject(const Name: string; const Object_: IDispatch;
  AddMembers: Boolean);
begin
  FGlobalObjects.AddOrSetValue(Name, Pointer(Object_));
  if Script.AddNamedItem(PChar(Name), SCRIPTITEM_ISVISIBLE {or SCRIPTITEM_ISPERSISTENT} or IfThen(AddMembers, SCRIPTITEM_GLOBALMEMBERS,0)) <> S_OK then
    RaiseLastOSError();
end;

constructor TActiveScript.Create(AOwner: TComponent);
begin
  inherited;
  FExceptionOnError := True;
  FLanguage := DefaultLanguage;
  FGlobalObjects := TDictionary<string, {IDispatch}Pointer>.Create();

  CreateEngine();
end;

procedure TActiveScript.CreateEngine;
var
  hr: HRESULT;
  ClsId: TGUID;
begin
  ScriptSite := TScriptSite.Create(Self);
  ScriptSite._AddRef();

  ClsId := CLSID_JScript; // Hardcode for now

  hr := CoCreateInstance(ClsId, nil, CLSCTX_INPROC_SERVER, IID_IActiveScript, Script);
  if hr <> S_OK then RaiseLastOSError();

  hr := Script.SetScriptSite(ScriptSite as IActiveScriptSite);
  if hr <> S_OK then RaiseLastOSError();

  ScriptParse := Script as IActiveScriptParse;
  hr := ScriptParse.InitNew();
  if hr <> S_OK then RaiseLastOSError();

end;

destructor TActiveScript.Destroy;
begin
  ReleaseEngine();
  FGlobalObjects.Free;
  inherited;
end;

function TActiveScript.Eval(const Expression: string): OleVariant;
var
  hr: HRESULT;
  ei: TExcepInfo;
begin
  HasErrorInfo := False;
  hr := ScriptParse.ParseScriptText(PChar(Expression), nil, nil, nil, 0, 0, SCRIPTTEXT_ISEXPRESSION, @Result, ei);
  if (hr <> S_OK) and (HasErrorInfo) and (FExceptionOnError) then
    raise EActiveScriptException.Create(ErrInfo.bstrDescription, ErrInfo.scode, ErrInfo.bstrSource, ErrInfo.bstrHelpFile, ErrInfo.dwHelpContext, ErrLnNum, ErrCharPos);
end;

procedure TActiveScript.ReleaseEngine;
begin
  Script.Close();
  ScriptParse := nil;
  Script := nil;

  ScriptSite._Release();
  ScriptSite := nil;
end;

procedure TActiveScript.Reset;
begin
  ReleaseEngine();
  FGlobalObjects.Clear();
  CreateEngine();
end;

procedure TActiveScript.SetLanguage(const Value: string);
begin
  if FLanguage <> Value then
  begin
    FLanguage := Value;
    // TODO: Find CLSID for Language
  end;
end;

{ EActiveScriptException }

constructor EActiveScriptException.Create(const Message: string;
  ErrorCode: HRESULT; const Source, HelpFile: string; HelpContext: Integer;
  LineNumber: Cardinal; CharacterPosition: Integer);
begin
  inherited Create(Message, ErrorCode, Source, HelpFile, HelpContext);
  FLineNumber := LineNumber;
  FCharacterPosition := CharacterPosition;
end;

initialization
  CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
finalization
  CoUninitialize();
end.
