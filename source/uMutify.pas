unit uMutify;

interface

uses
  SysUtils,
  Windows,
  ActiveX,
  ComObj;

const
  CLASS_IMMDeviceEnumerator: TGUID = '{BCDE0395-E52F-467C-8E3D-C4579291692E}';
  IID_IMMDeviceEnumerator: TGUID = '{A95664D2-9614-4F35-A746-DE8DB63617E6}';
  IID_IAudioEndpointVolume: TGUID = '{5CDF2C82-841E-4546-9722-0CF74078229A}';

type
  IAudioEndpointVolumeCallback = interface(IUnknown)
    ['{657804FA-D6AD-4496-8A60-352752AF4F89}']
  end;

  IAudioEndpointVolume = interface(IUnknown)
    ['{5CDF2C82-841E-4546-9722-0CF74078229A}']
    function RegisterControlChangeNotify(AudioEndPtVol
      : IAudioEndpointVolumeCallback): HRESULT; stdcall;
    function UnregisterControlChangeNotify(AudioEndPtVol
      : IAudioEndpointVolumeCallback): HRESULT; stdcall;
    function GetChannelCount(out PInteger): HRESULT; stdcall;
    function SetMasterVolumeLevel(fLevelDB: single; pguidEventContext: PGUID)
      : HRESULT; stdcall;
    function SetMasterVolumeLevelScalar(fLevelDB: single;
      pguidEventContext: PGUID): HRESULT; stdcall;
    function GetMasterVolumeLevel(out fLevelDB: single): HRESULT; stdcall;
    function GetMasterVolumeLevelScaler(out fLevelDB: single): HRESULT; stdcall;
    function SetChannelVolumeLevel(nChannel: Integer; fLevelDB: double;
      pguidEventContext: PGUID): HRESULT; stdcall;
    function SetChannelVolumeLevelScalar(nChannel: Integer; fLevelDB: double;
      pguidEventContext: PGUID): HRESULT; stdcall;
    function GetChannelVolumeLevel(nChannel: Integer; out fLevelDB: double)
      : HRESULT; stdcall;
    function GetChannelVolumeLevelScalar(nChannel: Integer; out fLevel: double)
      : HRESULT; stdcall;
    function SetMute(bMute: Boolean; pguidEventContext: PGUID)
      : HRESULT; stdcall;
    function GetMute(out bMute: Boolean): HRESULT; stdcall;
    function GetVolumeStepInfo(pnStep: Integer; out pnStepCount: Integer)
      : HRESULT; stdcall;
    function VolumeStepUp(pguidEventContext: PGUID): HRESULT; stdcall;
    function VolumeStepDown(pguidEventContext: PGUID): HRESULT; stdcall;
    function QueryHardwareSupport(out pdwHardwareSupportMask): HRESULT; stdcall;
    function GetVolumeRange(out pflVolumeMindB: double;
      out pflVolumeMaxdB: double; out pflVolumeIncrementdB: double)
      : HRESULT; stdcall;
  end;

  IAudioMeterInformation = interface(IUnknown)
    ['{C02216F6-8C67-4B5B-9D00-D008E73E0064}']
  end;

  IPropertyStore = interface(IUnknown)
  end;

  IMMDevice = interface(IUnknown)
    ['{D666063F-1587-4E43-81F1-B948E807363F}']
    function Activate(const refId: TGUID; dwClsCtx: DWORD;
      pActivationParams: PInteger; out pEndpointVolume: IAudioEndpointVolume)
      : HRESULT; stdCall;
    function OpenPropertyStore(stgmAccess: DWORD;
      out ppProperties: IPropertyStore): HRESULT; stdcall;
    function GetId(out ppstrId: PLPWSTR): HRESULT; stdcall;
    function GetState(out State: Integer): HRESULT; stdcall;
  end;

  IMMDeviceCollection = interface(IUnknown)
    ['{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}']
  end;

  IMMNotificationClient = interface(IUnknown)
    ['{7991EEC9-7E89-4D85-8390-6C703CEC60C0}']
  end;

  IMMDeviceEnumerator = interface(IUnknown)
    ['{A95664D2-9614-4F35-A746-DE8DB63617E6}']
    function EnumAudioEndpoints(dataFlow: TOleEnum; deviceState: SYSUINT;
      DevCollection: IMMDeviceCollection): HRESULT; stdcall;
    function GetDefaultAudioEndpoint(EDF: SYSUINT; ER: SYSUINT;
      out Dev: IMMDevice): HRESULT; stdcall;
    function GetDevice(pwstrId: pointer; out Dev: IMMDevice): HRESULT; stdcall;
    function RegisterEndpointNotificationCallback
      (pClient: IMMNotificationClient): HRESULT; stdcall;
  end;

procedure SetMasterVolume(fLevelDB: single);
function GetMasterVolume: single;

implementation

procedure SetMasterVolume(fLevelDB: single);
var
  pEndpointVolume: IAudioEndpointVolume;
  LDeviceEnumerator: IMMDeviceEnumerator;
  Dev: IMMDevice;
begin
  if not Succeeded(CoCreateInstance(CLASS_IMMDeviceEnumerator, nil,
    CLSCTX_INPROC_SERVER, IID_IMMDeviceEnumerator, LDeviceEnumerator)) then
    RaiseLastOSError;
  if not Succeeded(LDeviceEnumerator.GetDefaultAudioEndpoint($00000000,
    $00000000, Dev)) then
    RaiseLastOSError;

  if not Succeeded(Dev.Activate(IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER,
    nil, pEndpointVolume)) then
    RaiseLastOSError;

  if not Succeeded(pEndpointVolume.SetMasterVolumeLevelScalar(fLevelDB, nil))
  then
    RaiseLastOSError;
end;

function GetMasterVolume: single;
var
  pEndpointVolume: IAudioEndpointVolume;
  LDeviceEnumerator: IMMDeviceEnumerator;
  Dev: IMMDevice;
begin
  if not Succeeded(CoCreateInstance(CLASS_IMMDeviceEnumerator, nil,
    CLSCTX_INPROC_SERVER, IID_IMMDeviceEnumerator, LDeviceEnumerator)) then
    RaiseLastOSError;
  if not Succeeded(LDeviceEnumerator.GetDefaultAudioEndpoint($00000000,
    $00000000, Dev)) then
    RaiseLastOSError;

  if not Succeeded(Dev.Activate(IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER,
    nil, pEndpointVolume)) then
    RaiseLastOSError;

  if not Succeeded(pEndpointVolume.GetMasterVolumeLevelScaler(Result)) then
    RaiseLastOSError;
end;

initialization

CoInitialize(nil);

finalization

CoUninitialize;

end.
