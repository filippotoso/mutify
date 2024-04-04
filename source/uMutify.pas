unit uMutify;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Winapi.ActiveX;

type
  // Interface IPropertyStore
  IPropertyStore = interface(IUnknown)
    ['{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}']
    function GetCount(out cProps: DWORD): HResult; stdcall;

    function GetAt(iProp: DWORD; out pkey: PPROPERTYKEY): HResult; stdcall;

    function GetValue(const key: PROPERTYKEY; out pv: PROPVARIANT)
      : HResult; stdcall;

    function SetValue(const key: PROPERTYKEY; const propvar: PROPVARIANT)
      : HResult; stdcall;

    function Commit: HResult; stdcall;

  end;

  IID_IPropertyStore = IPropertyStore;


  // Interface IMMNotificationClient
  // ===============================
  {
    The IMMNotificationClient interface provides notifications when an audio endpoint device
    is added or removed, when the state or properties of an endpoint device change,
    or when there is a change in the default role assigned to an endpoint device.
    Unlike the other interfaces in this section, which are implemented by the MMDevice API system component,
    an MMDevice API client implements the IMMNotificationClient interface.
    To receive notifications, the client passes a pointer to its IMMNotificationClient interface instance
    as a parameter to the IMMDeviceEnumerator.RegisterEndpointNotificationCallback method.
  }

  IMMNotificationClient = interface(IUnknown)
    ['{7991EEC9-7E89-4D85-8390-6C703CEC60C0}']

    function OnDeviceStateChanged(pwstrDeviceId: LPCWSTR; dwNewState: DWORD)
      : HResult; stdcall;
    // Parameters
    // pwstrDeviceId [in]
    // Pointer to the endpoint ID string that identifies the audio endpoint device.
    // This parameter points to a null-terminated, wide-character string containing the endpoint ID.
    // The string remains valid for the duration of the call.
    // dwNewState [in]
    // Specifies the new state of the endpoint device.
    // The value of this parameter is one of the following DEVICE_STATE_XXX constants:
    // DEVICE_STATE_ACTIVE
    // DEVICE_STATE_DISABLED
    // DEVICE_STATE_NOTPRESENT
    // DEVICE_STATE_UNPLUGGED

    function OnDeviceAdded(pwstrDeviceId: LPCWSTR): HResult; stdcall;
    // The OnDeviceAdded method indicates that a new audio endpoint device has been added.
    // Parameters
    // pwstrDeviceId [in]
    // Pointer to the endpoint ID string that identifies the audio endpoint device.
    // This parameter points to a null-terminated, wide-character string containing
    // the endpoint ID. The string remains valid for the duration of the call.

    function OnDeviceRemoved(pwstrDeviceId: LPCWSTR): HResult; stdcall;
    // Parameters
    // flow [in]
    // The data-flow direction of the endpoint device.
    // This parameter is set to one of the following EDataFlow enumeration values:
    // - eRender
    // - eCapture
    // The data-flow direction for a rendering device is eRender.
    // The data-flow direction for a capture device is eCapture.
    // role [in]
    // The device role of the audio endpoint device.
    // This parameter is set to one of the following ERole enumeration values:
    // - eConsole
    // - eMultimedia
    // - eCommunications
    // pwstrDefaultDeviceId [in]
    // Pointer to the endpoint ID string that identifies the audio endpoint device.
    // This parameter points to a null-terminated, wide-character string containing the endpoint ID.
    // The string remains valid for the duration of the call.
    // If the user has removed or disabled the default device for a particular role,
    // and no other device is available to assume that role, then pwstrDefaultDeviceId is NULL.

    function OnDefaultDeviceChanged(flow: Cardinal; role: Cardinal;
      pwstrDefaultDeviceId: LPCWSTR): HResult; stdcall;
    // The OnDefaultDeviceChanged method notifies the client that the default audio endpoint device for a particular device role has changed.
    // Parameters
    // flow [in]
    // The data-flow direction of the endpoint device.
    // This parameter is set to one of the following EDataFlow enumeration values:
    // - eRender  : The data-flow direction for a rendering device is eRender.
    // - eCapture : The data-flow direction for a capture device is eCapture.
    // role [in]
    // The device role of the audio endpoint device.
    // This parameter is set to one of the following ERole enumeration values:
    // - eConsole
    // - eMultimedia
    // - eCommunications
    // pwstrDefaultDevice [in]
    // Pointer to the endpoint ID string that identifies the audio endpoint device.
    // This parameter points to a null-terminated, wide-character string containing the endpoint ID.
    // The string remains valid for the duration of the call.
    // If the user has removed or disabled the default device for a particular role,
    // and no other device is available to assume that role, then pwstrDefaultDevice is NIL.

    function OnPropertyValueChanged(pwstrDeviceId: LPCWSTR; key: PROPERTYKEY)
      : HResult; stdcall;
    // The OnPropertyValueChanged method indicates that the value of a property belonging to an audio endpoint device has changed.
    // Parameters
    // pwstrDeviceId [in]
    // Pointer to the endpoint ID string that identifies the audio endpoint device.
    // This parameter points to a null-terminated, wide-character string that contains the endpoint ID.
    // The string remains valid for the duration of the call.
    // key [in]
    // A PROPERTYKEY structure that specifies the property.
    // The structure contains the property-set GUID and an index identifying a property within the set.
    // The structure is passed by value. It remains valid for the duration of the call.
    // For more information about PROPERTYKEY, see the Windows SDK documentation.
    //
    // Remarks
    // A call to the IPropertyStore.SetValue method that successfully changes the value of
    // a property of an audio endpoint device generates a call to OnPropertyValueChanged.
    // For more information about IPropertyStore.SetValue, see the Windows SDK documentation.
    // A client can use the key parameter to retrieve the new property value.
  end;

  IID_IMMNotificationClient = IMMNotificationClient;

  // Interface IMMDevice
  // ===================
  {
    The IMMDevice interface encapsulates the generic features of a multimedia device resource.
    In the current implementation of the MMDevice API, the only type of device resource that an
    IMMDevice interface can represent is an audio endpoint device.

    NOTE:
    A client can obtain an IMMDevice interface from one of the following methods:
    IMMDeviceCollection.Item
    IMMDeviceEnumerator.GetDefaultAudioEndpoint
    IMMDeviceEnumerator.GetDevice
  }
  IMMDevice = interface(IUnknown)
    ['{D666063F-1587-4E43-81F1-B948E807363F}']

    function Activate(const iid: TGUID; dwClsCtx: DWORD;
      { In_opt } pActivationParams: PPROPVARIANT; out ppInterface): HResult;
      stdcall; // Replaced IUNKOWN pointer to a pointer, as described on ms-docs:
    // https://docs.microsoft.com/us-en/windows/win32/api/mmdeviceapi/nf-mmdeviceapi-immdevice-activate
    function OpenPropertyStore(stgmAccess: DWORD;
      out ppProperties: IPropertyStore): HResult; stdcall;

    function GetId(out ppstrId: PWideChar): HResult; stdcall;
    // 250815a, modified; issue reported by mbergstrand

    function GetState(out pdwState: UINT): HResult; stdcall;
    // Parameters
    // pdwState [out]
    // Pointer to a DWORD variable into which the method writes the current state of the device.
    // The device-state value is one of the following DEVICE_STATE_XXX constants:
    // DEVICE_STATE_ACTIVE
    // DEVICE_STATE_DISABLED
    // DEVICE_STATE_NOTPRESENT
    // DEVICE_STATE_UNPLUGGED

  end;

  IID_IMMDevice = IMMDevice;

  // Interface IMMDeviceCollection
  // =============================
  {
    The IMMDeviceCollection interface represents a collection of multimedia device resources.
    In the current implementation, the only device resources that the MMDevice API can create collections
    of all audio endpoint devices.
    A client can obtain a reference to an IMMDeviceCollection interface instance by calling the
    IMMDeviceEnumerator.EnumAudioEndpoints method.
    This method creates a collection of endpoint objects, each of which represents an
    audio endpoint device in the system.
    Each endpoint object in the collection supports the IMMDevice and IMMEndpoint interfaces.
  }
  IMMDeviceCollection = interface(IUnknown)
    ['{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}']

    function GetCount(out pcDevices: UINT): HResult; stdcall;

    function Item(nDevice: UINT; out ppDevice: IMMDevice): HResult; stdcall;

  end;

  IID_IMMDeviceCollection = IMMDeviceCollection;

  // Interface IMMDeviceEnumerator
  // =============================
  {
    The IMMDeviceEnumerator interface provides methods for enumerating multimedia device resources.
    In the current implementation of the MMDevice API, the only device resources that this interface
    can enumerate are audio endpoint devices.
    A client obtains a reference to an IMMDeviceEnumerator interface by calling the CoCreateInstance function,
    as described previously (see MMDevice API).

    NOTE:
    The device resources enumerated by the methods in the IMMDeviceEnumerator interface are represented
    as collections of objects with IMMDevice interfaces.
    A collection has an IMMDeviceCollection interface.
    The IMMDeviceEnumerator.EnumAudioEndpoints method creates a device collection.

    To obtain a pointer to the IMMDevice interface of an item in a device collection,
    the client calls the IMMDeviceCollection.Item method.
  }
  IMMDeviceEnumerator = interface(IUnknown)
    ['{A95664D2-9614-4F35-A746-DE8DB63617E6}']

    function EnumAudioEndpoints(dataFlow: Cardinal; const dwStateMask: Cardinal;
      // DEVICE_STATE_XXX constants
      out ppDevices: IMMDeviceCollection): HResult; stdcall;

    function GetDefaultAudioEndpoint(dataFlow: Cardinal; role: Cardinal;
      out ppEndpoint: IMMDevice): HResult; stdcall;

    function GetDevice(pwstrId: PWChar;
      // Pointer to a string containing the endpoint ID.
      // The caller typically obtains this string from the IMMDevice.GetId method or
      // from one of the methods in the IMMNotificationClient interface.
      out ppDevice: IMMDevice): HResult; stdcall;

    function RegisterEndpointNotificationCallback
      (pClient: IMMNotificationClient): HResult; stdcall;

    function UnregisterEndpointNotificationCallback
      (pClient: IMMNotificationClient): HResult; stdcall;

  end;

  IID_IMMDeviceEnumerator = IMMDeviceEnumerator;

  // Interface IAudioSessionEvents
  // =============================
  // Application initiated events.
  IAudioSessionEvents = interface(IUnknown)
    ['{24918ACC-64B3-37C1-8CA9-74A66E9957A8}']
    function OnDisplayNameChanged(NewDisplayName: LPCWSTR;
      const EventContext: TGUID): HResult; stdcall;
    // Description:
    //
    // Called when the display name of an AudioSession changes.
    //
    // Parameters:
    //
    // newDisplayName - [in]
    // The new display name for the Audio Session.
    // EventContext - [in]
    // Context passed to SetDisplayName routine.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //

    function OnIconPathChanged(NewIconPath: LPCWSTR; const EventContext: TGUID)
      : HResult; stdcall;
    // Description:
    //
    // Called when the icon path of an AudioSession changes.
    //
    // Parameters:
    //
    // NewIconPath - [in]
    // The new icon path for the Audio Session.
    // EventContext - [in]
    // Context passed to SetIconPath routine.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //

    function OnSimpleVolumeChanged(NewVolume: Single; NewMute: BOOL;
      const EventContext: TGUID): HResult; stdcall;
    // Description:
    //
    // Called when the simple volume of an AudioSession changes.
    //
    // Parameters:
    //
    // newVolume - [in]
    // The new volume for the AudioSession.
    // newMute - [in]
    // The new mute state for the AudioSession.
    // EventContext - [in]
    // Context passed to SetVolume routine.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //

    function OnChannelVolumeChanged(ChannelCount: UINT;
      NewChannelArray: PSingle; ChangedChannel: UINT; const EventContext: TGUID)
      : HResult; stdcall;
    // Description:
    //
    // Called when the channel volume of an AudioSession changes.
    //
    // Parameters:
    //
    // ChannelCount - [in]
    // The number of channels in the channel array.
    // NewChannelVolumeArray - [in]
    // An array containing the new channel volumes.
    // ChangedChannel - [in]
    // -1 if all channnels were changed, otherwise the channel volume which changed,
    // 0..ChannelCount-1
    // EventContext - [in]
    // Context passed to SetVolume routine.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure (ignored)
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //

    function OnGroupingParamChanged(NewGroupingParam: TGUID;
      const EventContext: TGUID): HResult; stdcall;
    // Description:
    // Called when the grouping param of an Audio Session changes.
    //
    // Parameters:
    // NewGroupingParam - [in]
    // The new gropuing param for the Audio Session.
    // EventContext - [in]
    // Context passed to SetGroupingParam routine.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //

    // System initiated events.
    // ==========================

    function OnStateChanged(NewState: Cardinal): HResult; stdcall;
    // Description:
    //
    // Called when the state of an AudioSession changes.
    //
    // Parameters:
    //
    // newState - [in]
    // The new state for the AudioSession.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //

    function OnSessionDisconnected(DisconnectReason: Cardinal)
      : HResult; stdcall;
    // Description:
    // Called when the audio session has been disconnected.
    //
    // Parameters:
    // DisconnectReason - [in]
    // The reason for the disconnection.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Please note: The caller of this function ignores all return
    // codes from this method.
    //
  end;

  // IAudioSessionEvents
  IID_IAudioSessionEvents = IAudioSessionEvents;

  // Interface IAudioSessionControl >= Vista
  // ==============================
  // Client interface that allows control over a AudioSession.
  //
  IAudioSessionControl = interface(IUnknown)
    ['{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}']
    function GetState(out pRetVal: UINT): HResult; stdcall;
    // Description:
    //
    // Retrieves the current AudioSession state.
    //
    // Parameters:
    //
    // pRetVal - [out]
    // The current AudioSession state, either
    // AudioSessionStateActive or AudioSessionStateInactive
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Remarks:
    //
    // If an AudioSession has audio streams currently opened on the AudioSession,
    // the AudioSession is considered active, otherwise it is inactive.
    //

    function GetDisplayName(out pRetVal: LPWSTR): HResult; stdcall;
    // pRetVal must be freed by CoTaskMemFree

    function SetDisplayName(Val: LPCWSTR; EventContext: TGUID)
      : HResult; stdcall;
    // Description:
    //
    // Sets or retrieves the current display name of the AudioSession.
    //
    // Parameters:
    //
    // Value - [in]
    // A string containing the current display name for the AudioSession.
    //
    // pRetVal - [out]
    // A string containing the current display name for the
    // AudioSession. Caller should free this with CoTaskMemFree.
    //
    // The DisplayName may be in the form of a shell resource spcification, in which case
    // the volume UI will extract the display name for the current language from
    // the specified path.
    //
    // The shell resource specification is of the form:
    // <path(including %environmentvariable%)>\<dll name>,<resource ID>
    // So:"%windir%\system32\shell32.dll,-240" is an example.
    // EventContext - [in]
    // Context passed to OnDisplayNameChanged routine, GUID_NULL if NULL.
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    // Remarks:
    // The application hosting the session controls the display name, if the application has NOT set the display name,
    // then this will return an empty string ("").
    //

    function GetIconPath(out pRetVal: LPWSTR): HResult; stdcall;
    // pRetVal must be freed by CoTaskMemFree

    function SetIconPath(Val: LPCWSTR; EventContext: TGUID): HResult; stdcall;
    // Description:
    //
    // Sets or retrieves an icon resource associated with the session.
    //
    // Parameters:
    //
    // Value - [in]
    // A string containing a shell resource specification used to retrieve
    // the icon for the Audio Session.  The volume UI will pass this string
    // to the ExtractIcon(Ex) API to extract the icon that is displayed
    //
    // The shell resource specification is of the form:
    // <path(including %environmentvariable%)>\<dll name>,<resource ID>
    // So:"%windir%\system32\shell32.dll,-240" is an example.
    //
    // pRetVal - [out]
    // A string containing the shell resource specification for the
    // Audio Session. Caller should free this with CoTaskMemFree.
    //
    // EventContext - [in]
    // Context passed to OnIconPathChanged routine.
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    // Remarks:
    // The application hosting the session controls the icon path, if the application has NOT set the icon path,
    // then this will return an empty string ("").
    //

    function GetGroupingParam(out pRetVal: TGUID): HResult; stdcall;

    function SetGroupingParam(const OverrideValue: TGUID;
      const EventContext: TGUID): HResult; stdcall;
    // Description:
    //
    // Gets or sets the current grouping param of the Audio Session.
    //
    // Parameters:
    //
    // GroupingParam - [in]
    // The GUID grouping param for the current Audio Session.
    // pRetVal - [out]
    // The GUID grouping param for the current Audio Session.
    // EventContext - [in]
    // Context passed to OnGroupingParamChanged routine.
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Remarks:
    // Normally the volume control application (sndvol.exe) will launch a separate slider
    // for each audio session.  If an application wishes to override this behavior, it can
    // set the session grouping param to an application defined GUID.  When the volume
    // control application sees two sessions with the same session control, it will only
    // display a single slider for those applications.
    //
    // As an example, normally, if you launch two copies of sndrec32.exe, then you will see
    // two volume control sliders in the windows volume control application. If sndrec32.exe
    // sets the grouping param, then the volume control will only show one slider, even though
    // there are two sessions.
    //
    // Please note that there are still two sessions, each with its own volume control, and those
    // volume controls may not have the same value. If this is the case, then it is the responsibility
    // of the application to ensure that the volume control on each session has the same value.
    //

    function RegisterAudioSessionNotification(NewNotifications
      : IAudioSessionEvents): HResult; stdcall;
    // Description:
    //
    // Add a notification callback to the list of AudioSession notification
    // callbacks.
    //
    // Parameters:
    //
    // NewNotifications - [in]
    // An object implementing the IAudioSessionEvents
    // interface.
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    //

    function UnregisterAudioSessionNotification(NewNotifications
      : IAudioSessionEvents): HResult; stdcall;
    // Description:
    //
    // Remove a notification callback to the list of AudioSession notification
    // callbacks.
    //
    // Parameters:
    //
    // NewNotifications - [in]
    // An object implementing the IAudioSessionEvents
    // interface.
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Remarks:
    // Please note: This function is a "finalizer". As such,
    // assuming that the NewNotification parameter has been
    // previously registered for notification, this function has
    // no valid failure modes.
    //
  end;

  IID_IAudioSessionControl = IAudioSessionControl;

  // Interface IAudioSessionControl2 >= Windows 7
  // ===============================
  // AudioSession Control Extended Interface
  //
  IAudioSessionControl2 = interface(IAudioSessionControl)
    ['{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}']
    function GetSessionIdentifier(out pRetVal: LPWSTR): HResult; stdcall;
    // Description:
    //
    // Retrieves the AudioSession ID.
    //
    // Parameters:
    //
    // pRetVal - [out]
    // A string containing the ID of the AudioSession.
    // Freed with CoTaskMemFree
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Remarks:
    // Each AudioSession has a unique identifier string associated with it.
    // This ID string represents the identifier of the AudioSession. It is NOT unique across all instances - if there are two
    // instances of the application playing, they will both have the same session identifier.
    //

    function GetSessionInstanceIdentifier(out pRetVal: LPWSTR)
      : HResult; stdcall;
    // Description:
    //
    // Retrieves the AudioSession instance ID.
    //
    // Parameters:
    //
    // pRetVal - [out]
    // A string containing the instance ID of the AudioSession.
    // Freed with CoTaskMemFree
    //
    // Return values:
    //
    // S_OK        Success
    // FAILURECODE Failure
    //
    // Remarks:
    //
    // Each AudioSession has a unique identifier string associated with it.
    // This ID string represents the particular instance of the AudioSession.
    //
    // The session instance identifier is unique across all instances, if there are two instances of the application playing,
    // they will have different instance identifiers.
    //

    function GetProcessId(out pRetVal: DWORD): HResult; stdcall;
    // Description:
    //
    // Retrieves the AudioSession Process ID.
    //
    // Parameters:
    //
    // pRetVal - [out]
    // A string containing the ID of the AudioSession.
    // Freed with CoTaskMemFree
    //
    // Return values:
    //
    // S_OK: Success
    // AUDCLNT_E_NO_SINGLE_PROCESS: The session spans more than one process.
    // In this case, pRetVal receives the initial identifier of the process that created the
    // session. To use this value, include the following definition:
    // const AUDCLNT_S_NO_SINGLE_PROCESS = $8890000D;
    // NOTE: This constant is defined in WinApi.CoreAudioApi.AudioClient.pas and is NOT present in any headerfile.
    // SEE: https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nf-audiopolicy-iaudiosessioncontrol2-getprocessid
    // FAILURECODE: Failure
    //
    // Remarks:
    //
    //

    function IsSystemSoundsSession(): HResult; stdcall;
    // Description:
    //
    // Determines if the specified session is the system sounds session
    //
    // Parameters:
    //
    // None
    //
    // Return values:
    //
    // S_OK        Success - The session is the system sounds session.
    // S_FALSE     Success - The session is NOT the system sounds session.
    // FAILURECODE Failure
    //

    function SetDuckingPreference(const optOut: BOOL): HResult; stdcall;
    // A BOOL variable that enables or disables system auto-ducking.
    // Description:
    //
    // Allows client to set its ducking preference
    //
    // Parameters:
    // optOut - [in]
    // Indicates whether caller wants to opt out of
    // system auto-ducking
    //
    // Remarks
    //
    // An application should call this method in advance of receiving
    // a ducking notification (generally at stream create time).  This
    // method can be called dynamically (i.e. over and over) as its
    // desire to opt in or opt out of ducking changes.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
  end;

  // IAudioSessionControl2
  IID_IAudioSessionControl2 = IAudioSessionControl2;

  // Interface ISimpleAudioVolume
  // ============================
  //
  ISimpleAudioVolume = interface(IUnknown)
    ['{87CE5498-68D6-44E5-9215-6DA47EF883D8}']
    function SetMasterVolume(fLevel: Single; EventContext: TGUID)
      : HResult; stdcall;
    // Description:
    //
    // Set the master volume of the current audio client.
    //
    // Parameters:
    //
    // fLevel - [in]
    // New amplitude for the audio stream.
    // EventContext - [in]
    // Context passed to notification routine, GUID_NULL if NULL.
    //
    // See Also:
    //
    // ISimpleAudioVolume.GetMasterVolume
    //
    // Return values:
    //
    // S_OK        Successful completion.
    //
    //

    function GetMasterVolume(out pfLevel: Single): HResult; stdcall;
    // Description:
    //
    // Get the master volume of the current audio client.
    //
    // Parameters:
    //
    // pfLevel - [out]
    // New amplitude for the audio stream.
    //
    // See Also:
    //
    // ISimpleAudioVolume.SetMasterVolume
    //
    // Return values:
    //
    // S_OK        Successful completion.
    // OTHER       Other error.
    //
    //

    function SetMute(bMute: BOOL; EventContext: TGUID): HResult; stdcall;
    // Description:
    //
    // Set the mute state of the current audio client.
    //
    // Parameters:
    //
    // bMute - [in]
    // New mute for the audio stream.
    // EventContext - [in]
    // Context passed to notification routine, GUID_NULL if NULL.
    //
    // See Also:
    //
    // ISimpleAudioVolume.SetMute
    //
    // Return values:
    //
    // S_OK        Successful completion.
    // OTHER       Other error.
    //

    function GetMute(out pbMute: BOOL): HResult; stdcall;
    // Description:
    //
    // Get the mute state of the current audio client.
    //
    // Parameters:
    //
    // bMute - [out]
    // Current mute for the audio stream.
    //
    // See Also:
    //
    // ISimpleAudioVolume.GetMute
    //
    // Return values:
    //
    // S_OK        Successful completion.
    // OTHER       Other error.
    //
    //
  end;

  IID_ISimpleAudioVolume = ISimpleAudioVolume;

  // Interface IAudioSessionManager
  // ==============================
  // Notification on session changes.
  //
  IAudioSessionManager = interface(IUnknown)
    ['{BFA971F1-4D5E-40BB-935E-967039BFBEE4}']
    function GetAudioSessionControl(AudioSessionGuid: TGUID;
      // may be @GUID_NULL or Nil when nothing is specified.
      StreamFlag: UINT; out SessionControl: IAudioSessionControl)
      : HResult; stdcall;
    // Description:
    //
    // Return an audio session control for the current process.
    //
    // Parameters:
    // AudioSessionGuid - [in]
    // Session ID for the session.
    // StreamFlags    - [in]
    // Combination of AUDCLNT_STREAMFLAGS_XYZ flags, defined MfPack.AudioSessionTypes
    // SessionControl   - [out]
    // Returns a pointer to an audio session control for the current process.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

    function GetSimpleAudioVolume(AudioSessionGuid: TGUID;
      // may be @GUID_NULL or Nil when nothing is specified.
      StreamFlag: UINT; out AudioVolume: ISimpleAudioVolume): HResult; stdcall;
    // Description:
    //
    // Return an audio volume control for the current process.
    //
    // Parameters:
    // AudioSessionGuid - [in]
    // Session ID for the session.
    // StreamFlags - [in]
    // Combination of AUDCLNT_STREAMFLAGS_XYZ flags, defined MfPack.AudioSessionTypes
    // AudioVolume - [out]
    // Returns a pointer to an audio volume control for a session in the current process.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

  end;

  IID_IAudioSessionManager = IAudioSessionManager;

  // Interface IAudioVolumeDuckNotification >= Windows 7
  // ======================================
  // Notification on session changes.
  //
  IAudioVolumeDuckNotification = interface(IUnknown)
    ['{C3B284D4-6D39-4359-B3CF-B56DDB3BB39C}']
    function OnVolumeDuckNotification(sessionID: LPCWSTR;
      countCommunicationSessions: UINT32): HResult; stdcall;
    // Description:
    //
    // Notification of a pending system auto-duck
    //
    // Parameters:
    // sessionID - [in]
    // Session instance ID of the communications session
    // creating the auto-ducking event.
    // countCommunicationsSessions - [in]
    // the number of active communications sessions (first
    // session is 1).
    //
    // Remarks
    // If applications wish to opt out of ducking, they must call
    // IAudioVolumeDuck.SetDuckingPreference()
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

    function OnVolumeUnduckNotification(sessionID: LPCWSTR): HResult; stdcall;
    // Description:
    //
    // Notification of a pending system auto-unduck
    //
    // Parameters:
    // sessionID - [in]
    // Session instance ID of the communications session
    // that is terminating.
    // countCommunicationsSessions - [in]
    // the number of active communications sessions (last
    // session is 1).
    //
    // Remarks
    // This is simply a notification that
    // the communications stream that initiated the ducking is terminating.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //
  end;

  IID_IAudioVolumeDuckNotification = IAudioVolumeDuckNotification;

  // Interface IAudioSessionNotification
  // ===================================
  // Audio Session Notification Interface
  //
  IAudioSessionNotification = interface(IUnknown)
    ['{641DD20B-4D41-49CC-ABA3-174B9477BB08}']

    function OnSessionCreated(NewSession: IAudioSessionControl)
      : HResult; stdcall;

  end;

  IID_IAudioSessionNotification = IAudioSessionNotification;

  // Interface IAudioSessionEnumerator
  // =================================
  // Audio Session Enumerator Interface
  //
  IAudioSessionEnumerator = interface(IUnknown)
    ['{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}']

    function GetCount(out SessionCount: integer): HResult; stdcall;

    function GetSession(SessionCount: integer;
      out Session: IAudioSessionControl): HResult; stdcall;

  end;

  IID_IAudioSessionEnumerator = IAudioSessionEnumerator;

  // Interface IAudioSessionManager2 >= Windows 7
  // ===============================
  // Manage interface for all submixes - Supports
  // enumeration and notification of submixes.
  // Also provides support for ducking notifications.
  //
  IAudioSessionManager2 = interface(IAudioSessionManager)
    ['{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}']
    function GetSessionEnumerator(out SessionEnum: IAudioSessionEnumerator)
      : HResult; stdcall;
    // Description:
    //
    // Return the audio session enumerator.
    //
    // Parameters:
    // SessionList - [out]
    // An IAudioSessionEnumerator interface that
    // can enumerate the audio sessions.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

    function RegisterSessionNotification(SessionNotification
      : IAudioSessionNotification): HResult; stdcall;
    // Description:
    //
    // Add the specified process ID as a target for session
    // notifications.
    //
    // Parameters:
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

    function UnregisterSessionNotification(SessionNotification
      : IAudioSessionNotification): HResult; stdcall;
    // Description:
    //
    // Remove the specified process ID as a target for session
    // notifications.
    //
    // Parameters:
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

    function RegisterDuckNotification(sessionID: LPCWSTR;
      duckNotification: IAudioVolumeDuckNotification): HResult; stdcall;
    // Description:
    //
    // Register for notification of a pending system auto-duck
    //
    // Parameters:
    // sessionID - [in]
    // A filter.  Applications like media players
    // that are interested in their sessions will
    // pass their own session instance ID.  Other applications
    // that want to see all the ducking notifications
    // can pass NULL.
    // duckNotification - [in]
    // Object which implements the
    // IAudioVolumeDuckNotification interface
    // which will receive new notifications.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
    //

    function UnregisterDuckNotification(duckNotification
      : IAudioVolumeDuckNotification): HResult; stdcall;
    // Description:
    //
    // Unregisters for notification of a pending system auto-duck
    //
    // Parameters:
    // duckNotification - [in]
    // Object which implements the
    // IAudioVolumeDuckNotification interface
    // previously registered, which will be
    // no longer receive notifications.  Please
    // note that after this routine returns, no
    // all pending notifications will have been
    // processed.
    //
    // Return values:
    // S_OK        Success
    // FAILURECODE Failure
  end;

  IID_IAudioSessionManager2 = IAudioSessionManager2;

const
  CLSID_MMDeviceEnumerator: TGUID = '{BCDE0395-E52F-467C-8E3D-C4579291692E}';

implementation

end.
