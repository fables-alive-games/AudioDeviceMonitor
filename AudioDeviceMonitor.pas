unit AudioDeviceMonitor;

interface

uses
  System.Classes, Winapi.Windows, Winapi.ActiveX, System.Win.ComObj, 
  Winapi.MMSystem, System.SysUtils, Vcl.Forms, System.Threading;

const
  CLSID_MMDeviceEnumerator: TGUID = '{BCDE0395-E52F-467C-8E3D-C4579291692E}';
  IID_IMMDeviceEnumerator: TGUID = '{A95664D2-9614-4F35-A746-DE8DB63617E6}';
  IID_IMMNotificationClient: TGUID = '{7991EEC9-7E89-4D85-8390-6C703CEC60C0}';
  
  // Property Keys
  PKEY_Device_FriendlyName: PROPERTYKEY = (fmtid: '{A45C254E-DF1C-4EFD-8020-67D146A850E0}'; pid: 14);
  PKEY_Device_DeviceDesc: PROPERTYKEY = (fmtid: '{A45C254E-DF1C-4EFD-8020-67D146A850E0}'; pid: 2);
  PKEY_Device_InterfaceFriendlyName: PROPERTYKEY = (fmtid: '{026E516E-B814-414B-83CD-856D6FEF4822}'; pid: 2);

  // Device States
  DEVICE_STATE_ACTIVE = $00000001;
  DEVICE_STATE_DISABLED = $00000002;
  DEVICE_STATE_NOTPRESENT = $00000004;
  DEVICE_STATE_UNPLUGGED = $00000008;
  DEVICE_STATEMASK_ALL = $0000000F;

  // PropVariant Types
  VT_EMPTY = 0;
  VT_NULL = 1;
  VT_I2 = 2;
  VT_I4 = 3;
  VT_R4 = 4;
  VT_R8 = 5;
  VT_CY = 6;
  VT_DATE = 7;
  VT_BSTR = 8;
  VT_DISPATCH = 9;
  VT_ERROR = 10;
  VT_BOOL = 11;
  VT_VARIANT = 12;
  VT_UNKNOWN = 13;
  VT_DECIMAL = 14;
  VT_I1 = 16;
  VT_UI1 = 17;
  VT_UI2 = 18;
  VT_UI4 = 19;
  VT_I8 = 20;
  VT_UI8 = 21;
  VT_INT = 22;
  VT_UINT = 23;
  VT_VOID = 24;
  VT_HRESULT = 25;
  VT_PTR = 26;
  VT_SAFEARRAY = 27;
  VT_CARRAY = 28;
  VT_USERDEFINED = 29;
  VT_LPSTR = 30;
  VT_LPWSTR = 31;

type
  EDataFlow = (
    eRender,
    eCapture,
    eAll,
    EDataFlow_enum_count
  );

  ERole = (
    eConsole,
    eMultimedia,
    eCommunications,
    ERole_enum_count
  );

  PROPERTYKEY = record
    fmtid: TGUID;
    pid: DWORD;
  end;

  PROPVARIANT = record
    vt: Word;
    wReserved1, wReserved2, wReserved3: Word;
    case Integer of
      0: (cVal: Char);
      1: (bVal: Byte);
      2: (iVal: Smallint);
      3: (uiVal: Word);
      4: (lVal: Longint);
      5: (ulVal: LongWord);
      6: (intVal: Integer);
      7: (uintVal: Cardinal);
      8: (hVal: Largeint);
      9: (uhVal: Int64);
      10: (fltVal: Single);
      11: (dblVal: Double);
      12: (boolVal: WordBool);
      13: (scode: LongInt);
      14: (cyVal: Currency);
      15: (date: TDateTime);
      16: (bstrVal: PWideChar);
      17: (punkVal: Pointer);
      18: (pdispVal: Pointer);
      19: (pszVal: PAnsiChar);
      20: (pwszVal: PWideChar);
  end;

  TAudioDeviceInfo = record
    DeviceId: string;
    DeviceName: string;
    Flow: EDataFlow;
    Role: ERole;
    State: DWORD;
    IsDefault: Boolean;
  end;

  IPropertyStore = interface(IUnknown)
    ['{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}']
    function GetCount(out cProps: DWORD): HResult; stdcall;
    function GetAt(iProp: DWORD; out pkey: PROPERTYKEY): HResult; stdcall;
    function GetValue(const key: PROPERTYKEY; var pv: PROPVARIANT): HResult; stdcall;
    function SetValue(const key: PROPERTYKEY; const propvar: PROPVARIANT): HResult; stdcall;
    function Commit: HResult; stdcall;
  end;

  IMMNotificationClient = interface(IUnknown)
    ['{7991EEC9-7E89-4D85-8390-6C703CEC60C0}']
    function OnDeviceStateChanged(const pwstrDeviceId: PWideChar; dwNewState: DWORD): HResult; stdcall;
    function OnDeviceAdded(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceRemoved(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDefaultDeviceChanged(flow: EDataFlow; role: ERole; const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnPropertyValueChanged(const pwstrDeviceId: PWideChar; const key: PROPERTYKEY): HResult; stdcall;
  end;

  IMMDevice = interface(IUnknown)
    ['{D666063F-1587-4E43-81F1-B948E807363F}']
    function Activate(const iid: TGUID; dwClsCtx: DWORD; pActivationParams: Pointer; out ppInterface: Pointer): HResult; stdcall;
    function OpenPropertyStore(stgmAccess: DWORD; out ppProperties: IPropertyStore): HResult; stdcall;
    function GetId(out ppstrId: PWideChar): HResult; stdcall;
    function GetState(out pdwState: DWORD): HResult; stdcall;
  end;

  IMMDeviceCollection = interface(IUnknown)
    ['{0BD7A1BE-7A1A-44DB-8397-C0A0B755BFA8}']
    function GetCount(out pcDevices: UINT): HResult; stdcall;
    function Item(nDevice: UINT; out ppDevice: IMMDevice): HResult; stdcall;
  end;

  IMMDeviceEnumerator = interface(IUnknown)
    ['{A95664D2-9614-4F35-A746-DE8DB63617E6}']
    function EnumAudioEndpoints(dataFlow: EDataFlow; dwStateMask: DWORD; out ppDevices: IMMDeviceCollection): HResult; stdcall;
    function GetDefaultAudioEndpoint(dataFlow: EDataFlow; role: ERole; out ppEndpoint: IMMDevice): HResult; stdcall;
    function GetDevice(pwstrId: PWideChar; out ppDevice: IMMDevice): HResult; stdcall;
    function RegisterEndpointNotificationCallback(pClient: IMMNotificationClient): HResult; stdcall;
    function UnregisterEndpointNotificationCallback(pClient: IMMNotificationClient): HResult; stdcall;
  end;

  // Event Types
  TAudioDeviceChangedEvent = procedure(Sender: TObject; const DeviceInfo: TAudioDeviceInfo) of object;
  TAudioDeviceSimpleEvent = procedure(Sender: TObject) of object;

  TAudioDeviceMonitor = class(TComponent, IMMNotificationClient)
  private
    FDeviceEnumerator: IMMDeviceEnumerator;
    FInitialized: Boolean;
    FLastError: string;
    
    // Events
    FOnDefaultDeviceChanged: TAudioDeviceChangedEvent;
    FOnDeviceAdded: TAudioDeviceChangedEvent;
    FOnDeviceRemoved: TAudioDeviceChangedEvent;
    FOnDeviceStateChanged: TAudioDeviceChangedEvent;
    FOnPropertyValueChanged: TAudioDeviceChangedEvent;
    FOnAnyDeviceChanged: TAudioDeviceSimpleEvent;
    
    // Internal methods
    function GetDeviceFriendlyName(const DeviceId: PWideChar): string;
    function GetDeviceInfo(const DeviceId: PWideChar; Flow: EDataFlow; Role: ERole; State: DWORD): TAudioDeviceInfo;
    procedure SafeCallEvent(EventProc: TProc);
    function PropVariantClear(var pvar: PROPVARIANT): HResult;
    
  protected
    // IMMNotificationClient implementations
    function OnDefaultDeviceChanged(flow: EDataFlow; role: ERole; const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceAdded(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceRemoved(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceStateChanged(const pwstrDeviceId: PWideChar; dwNewState: DWORD): HResult; stdcall;
    function OnPropertyValueChanged(const pwstrDeviceId: PWideChar; const key: PROPERTYKEY): HResult; stdcall;
    
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    
    // Public methods
    function Initialize: Boolean;
    function GetCurrentDefaultDevice(Flow: EDataFlow; Role: ERole): TAudioDeviceInfo;
    function GetAllAudioDevices: TArray<TAudioDeviceInfo>;
    function IsDeviceActive(const DeviceId: string): Boolean;
    
    // Properties
    property Initialized: Boolean read FInitialized;
    property LastError: string read FLastError;
    
  published
    // Published events
    property OnDefaultDeviceChanged: TAudioDeviceChangedEvent read FOnDefaultDeviceChanged write FOnDefaultDeviceChanged;
    property OnDeviceAdded: TAudioDeviceChangedEvent read FOnDeviceAdded write FOnDeviceAdded;
    property OnDeviceRemoved: TAudioDeviceChangedEvent read FOnDeviceRemoved write FOnDeviceRemoved;
    property OnDeviceStateChanged: TAudioDeviceChangedEvent read FOnDeviceStateChanged write FOnDeviceStateChanged;
    property OnPropertyValueChanged: TAudioDeviceChangedEvent read FOnPropertyValueChanged write FOnPropertyValueChanged;
    property OnAnyDeviceChanged: TAudioDeviceSimpleEvent read FOnAnyDeviceChanged write FOnAnyDeviceChanged;
  end;

// Helper functions
function DataFlowToString(Flow: EDataFlow): string;
function RoleToString(Role: ERole): string;
function DeviceStateToString(State: DWORD): string;
procedure Register;

implementation

uses
  Winapi.Ole2, Winapi.PropSys;

procedure Register;
begin
  RegisterComponents('Audio', [TAudioDeviceMonitor]);
end;

// Helper functions implementation
function DataFlowToString(Flow: EDataFlow): string;
begin
  case Flow of
    eRender: Result := 'Render (Playback)';
    eCapture: Result := 'Capture (Recording)';
    eAll: Result := 'All';
  else
    Result := 'Unknown';
  end;
end;

function RoleToString(Role: ERole): string;
begin
  case Role of
    eConsole: Result := 'Console';
    eMultimedia: Result := 'Multimedia';
    eCommunications: Result := 'Communications';
  else
    Result := 'Unknown';
  end;
end;

function DeviceStateToString(State: DWORD): string;
begin
  case State of
    DEVICE_STATE_ACTIVE: Result := 'Active';
    DEVICE_STATE_DISABLED: Result := 'Disabled';
    DEVICE_STATE_NOTPRESENT: Result := 'Not Present';
    DEVICE_STATE_UNPLUGGED: Result := 'Unplugged';
  else
    Result := 'Unknown';
  end;
end;

{ TAudioDeviceMonitor }

constructor TAudioDeviceMonitor.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FInitialized := False;
  FLastError := '';
  
  if not (csDesigning in ComponentState) then
    Initialize;
end;

destructor TAudioDeviceMonitor.Destroy;
begin
  try
    if FInitialized and Assigned(FDeviceEnumerator) then
    begin
      try
        FDeviceEnumerator.UnregisterEndpointNotificationCallback(Self as IMMNotificationClient);
      except
        // Ignore errors during cleanup
      end;
    end;
  finally
    try
      CoUninitialize;
    except
      // Ignore errors during cleanup
    end;
    inherited Destroy;
  end;
end;

function TAudioDeviceMonitor.Initialize: Boolean;
var
  hr: HResult;
begin
  Result := False;
  FLastError := '';
  
  try
    // Initialize COM
    hr := CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
    if FAILED(hr) and (hr <> RPC_E_CHANGED_MODE) then
    begin
      FLastError := Format('CoInitializeEx failed: 0x%x', [hr]);
      Exit;
    end;
    
    // Create MMDeviceEnumerator
    hr := CoCreateInstance(CLSID_MMDeviceEnumerator, nil, CLSCTX_INPROC_SERVER, 
                          IID_IMMDeviceEnumerator, FDeviceEnumerator);
    if FAILED(hr) then
    begin
      FLastError := Format('Failed to create MMDeviceEnumerator: 0x%x', [hr]);
      Exit;
    end;
    
    // Register notification callback
    if Assigned(FDeviceEnumerator) then
    begin
      hr := FDeviceEnumerator.RegisterEndpointNotificationCallback(Self as IMMNotificationClient);
      if FAILED(hr) then
      begin
        FLastError := Format('Failed to register notification callback: 0x%x', [hr]);
        Exit;
      end;
    end;
    
    FInitialized := True;
    Result := True;
    
  except
    on E: Exception do
    begin
      FLastError := 'Exception during initialization: ' + E.Message;
    end;
  end;
end;

function TAudioDeviceMonitor.PropVariantClear(var pvar: PROPVARIANT): HResult;
begin
  // Simple PropVariant cleanup - for more complex types, use Ole32.PropVariantClear
  if pvar.vt = VT_LPWSTR then
  begin
    if Assigned(pvar.pwszVal) then
    begin
      CoTaskMemFree(pvar.pwszVal);
      pvar.pwszVal := nil;
    end;
  end;
  pvar.vt := VT_EMPTY;
  Result := S_OK;
end;

function TAudioDeviceMonitor.GetDeviceFriendlyName(const DeviceId: PWideChar): string;
var
  Device: IMMDevice;
  PropertyStore: IPropertyStore;
  PropVar: PROPVARIANT;
  hr: HResult;
begin
  Result := 'Unknown Device';
  
  if not Assigned(FDeviceEnumerator) then Exit;
  
  try
    hr := FDeviceEnumerator.GetDevice(DeviceId, Device);
    if SUCCEEDED(hr) and Assigned(Device) then
    begin
      hr := Device.OpenPropertyStore(STGM_READ, PropertyStore);
      if SUCCEEDED(hr) and Assigned(PropertyStore) then
      begin
        FillChar(PropVar, SizeOf(PropVar), 0);
        hr := PropertyStore.GetValue(PKEY_Device_FriendlyName, PropVar);
        if SUCCEEDED(hr) and (PropVar.vt = VT_LPWSTR) and Assigned(PropVar.pwszVal) then
        begin
          Result := string(PropVar.pwszVal);
        end;
        PropVariantClear(PropVar);
      end;
    end;
  except
    on E: Exception do
      Result := 'Error: ' + E.Message;
  end;
end;

function TAudioDeviceMonitor.GetDeviceInfo(const DeviceId: PWideChar; Flow: EDataFlow; 
  Role: ERole; State: DWORD): TAudioDeviceInfo;
begin
  FillChar(Result, SizeOf(Result), 0);
  
  if Assigned(DeviceId) then
    Result.DeviceId := string(DeviceId)
  else
    Result.DeviceId := '';
    
  Result.DeviceName := GetDeviceFriendlyName(DeviceId);
  Result.Flow := Flow;
  Result.Role := Role;
  Result.State := State;
  Result.IsDefault := True; // This is called when device becomes default
end;

procedure TAudioDeviceMonitor.SafeCallEvent(EventProc: TProc);
begin
  if Assigned(EventProc) then
  begin
    try
      // Call event in main thread to avoid UI issues
      TThread.Queue(nil, EventProc);
    except
      on E: Exception do
      begin
        // Log error but don't propagate to prevent COM issues
        FLastError := 'Event callback error: ' + E.Message;
      end;
    end;
  end;
end;

function TAudioDeviceMonitor.OnDefaultDeviceChanged(flow: EDataFlow; role: ERole; 
  const pwstrDeviceId: PWideChar): HResult;
var
  DeviceInfo: TAudioDeviceInfo;
begin
  try
    DeviceInfo := GetDeviceInfo(pwstrDeviceId, flow, role, DEVICE_STATE_ACTIVE);
    
    SafeCallEvent(
      procedure
      begin
        if Assigned(FOnDefaultDeviceChanged) then
          FOnDefaultDeviceChanged(Self, DeviceInfo);
        if Assigned(FOnAnyDeviceChanged) then
          FOnAnyDeviceChanged(Self);
      end
    );
    
    Result := S_OK;
  except
    Result := E_FAIL;
  end;
end;

function TAudioDeviceMonitor.OnDeviceAdded(const pwstrDeviceId: PWideChar): HResult;
var
  DeviceInfo: TAudioDeviceInfo;
begin
  try
    DeviceInfo := GetDeviceInfo(pwstrDeviceId, eAll, eConsole, DEVICE_STATE_ACTIVE);
    
    SafeCallEvent(
      procedure
      begin
        if Assigned(FOnDeviceAdded) then
          FOnDeviceAdded(Self, DeviceInfo);
        if Assigned(FOnAnyDeviceChanged) then
          FOnAnyDeviceChanged(Self);
      end
    );
    
    Result := S_OK;
  except
    Result := E_FAIL;
  end;
end;

function TAudioDeviceMonitor.OnDeviceRemoved(const pwstrDeviceId: PWideChar): HResult;
var
  DeviceInfo: TAudioDeviceInfo;
begin
  try
    DeviceInfo := GetDeviceInfo(pwstrDeviceId, eAll, eConsole, DEVICE_STATE_NOTPRESENT);
    
    SafeCallEvent(
      procedure
      begin
        if Assigned(FOnDeviceRemoved) then
          FOnDeviceRemoved(Self, DeviceInfo);
        if Assigned(FOnAnyDeviceChanged) then
          FOnAnyDeviceChanged(Self);
      end
    );
    
    Result := S_OK;
  except
    Result := E_FAIL;
  end;
end;

function TAudioDeviceMonitor.OnDeviceStateChanged(const pwstrDeviceId: PWideChar; 
  dwNewState: DWORD): HResult;
var
  DeviceInfo: TAudioDeviceInfo;
begin
  try
    DeviceInfo := GetDeviceInfo(pwstrDeviceId, eAll, eConsole, dwNewState);
    
    SafeCallEvent(
      procedure
      begin
        if Assigned(FOnDeviceStateChanged) then
          FOnDeviceStateChanged(Self, DeviceInfo);
        if Assigned(FOnAnyDeviceChanged) then
          FOnAnyDeviceChanged(Self);
      end
    );
    
    Result := S_OK;
  except
    Result := E_FAIL;
  end;
end;

function TAudioDeviceMonitor.OnPropertyValueChanged(const pwstrDeviceId: PWideChar; 
  const key: PROPERTYKEY): HResult;
var
  DeviceInfo: TAudioDeviceInfo;
begin
  try
    DeviceInfo := GetDeviceInfo(pwstrDeviceId, eAll, eConsole, DEVICE_STATE_ACTIVE);
    
    SafeCallEvent(
      procedure
      begin
        if Assigned(FOnPropertyValueChanged) then
          FOnPropertyValueChanged(Self, DeviceInfo);
      end
    );
    
    Result := S_OK;
  except
    Result := E_FAIL;
  end;
end;

function TAudioDeviceMonitor.GetCurrentDefaultDevice(Flow: EDataFlow; Role: ERole): TAudioDeviceInfo;
var
  Device: IMMDevice;
  DeviceId: PWideChar;
  State: DWORD;
  hr: HResult;
begin
  FillChar(Result, SizeOf(Result), 0);
  
  if not FInitialized or not Assigned(FDeviceEnumerator) then Exit;
  
  try
    hr := FDeviceEnumerator.GetDefaultAudioEndpoint(Flow, Role, Device);
    if SUCCEEDED(hr) and Assigned(Device) then
    begin
      hr := Device.GetId(DeviceId);
      if SUCCEEDED(hr) then
      begin
        hr := Device.GetState(State);
        if SUCCEEDED(hr) then
        begin
          Result := GetDeviceInfo(DeviceId, Flow, Role, State);
          Result.IsDefault := True;
        end;
        CoTaskMemFree(DeviceId);
      end;
    end;
  except
    on E: Exception do
      FLastError := 'GetCurrentDefaultDevice error: ' + E.Message;
  end;
end;

function TAudioDeviceMonitor.GetAllAudioDevices: TArray<TAudioDeviceInfo>;
var
  DeviceCollection: IMMDeviceCollection;
  DeviceCount: UINT;
  Device: IMMDevice;
  DeviceId: PWideChar;
  State: DWORD;
  DeviceInfo: TAudioDeviceInfo;
  i: UINT;
  hr: HResult;
begin
  SetLength(Result, 0);
  
  if not FInitialized or not Assigned(FDeviceEnumerator) then Exit;
  
  try
    hr := FDeviceEnumerator.EnumAudioEndpoints(eAll, DEVICE_STATEMASK_ALL, DeviceCollection);
    if SUCCEEDED(hr) and Assigned(DeviceCollection) then
    begin
      hr := DeviceCollection.GetCount(DeviceCount);
      if SUCCEEDED(hr) then
      begin
        SetLength(Result, DeviceCount);
        for i := 0 to DeviceCount - 1 do
        begin
          hr := DeviceCollection.Item(i, Device);
          if SUCCEEDED(hr) and Assigned(Device) then
          begin
            hr := Device.GetId(DeviceId);
            if SUCCEEDED(hr) then
            begin
              hr := Device.GetState(State);
              if SUCCEEDED(hr) then
              begin
                DeviceInfo := GetDeviceInfo(DeviceId, eAll, eConsole, State);
                DeviceInfo.IsDefault := False; // Will be updated if needed
                Result[i] := DeviceInfo;
              end;
              CoTaskMemFree(DeviceId);
            end;
          end;
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      FLastError := 'GetAllAudioDevices error: ' + E.Message;
      SetLength(Result, 0);
    end;
  end;
end;

function TAudioDeviceMonitor.IsDeviceActive(const DeviceId: string): Boolean;
var
  Device: IMMDevice;
  State: DWORD;
  hr: HResult;
begin
  Result := False;
  
  if not FInitialized or not Assigned(FDeviceEnumerator) or (DeviceId = '') then Exit;
  
  try
    hr := FDeviceEnumerator.GetDevice(PWideChar(DeviceId), Device);
    if SUCCEEDED(hr) and Assigned(Device) then
    begin
      hr := Device.GetState(State);
      if SUCCEEDED(hr) then
        Result := (State = DEVICE_STATE_ACTIVE);
    end;
  except
    on E: Exception do
      FLastError := 'IsDeviceActive error: ' + E.Message;
  end;
end;

end.
