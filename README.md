# AudioDeviceMonitor

Delphi component for monitoring audio device changes using Windows Core Audio API. Detects changes like switching from PC speakers to headphones.

## Enhanced Version - What's New

### Error Handling
- HRESULT checks for all COM calls
- Exception handling and error logging
- LastError property for error tracking
- Initialize function for safe startup

### Thread Safety
- UI events run on main thread (TThread.Queue)
- Crash protection in COM callbacks
- Safe event calling mechanism

### Enhanced Event System
- **OnDefaultDeviceChanged**: Default device switched
- **OnDeviceAdded**: New device added
- **OnDeviceRemoved**: Device removed  
- **OnDeviceStateChanged**: Device state changed
- **OnPropertyValueChanged**: Device properties changed
- **OnAnyDeviceChanged**: Any device change

### Detailed Device Information
- TAudioDeviceInfo struct with rich device info
- Device friendly name retrieval
- Device ID, Flow, Role, State information
- Default device detection

### New Public Methods
- **GetCurrentDefaultDevice()**: Get current default device
- **GetAllAudioDevices()**: List all audio devices  
- **IsDeviceActive()**: Check device active status
- **Initialize()**: Manual initialization

### Helper Functions
- DataFlowToString(): Convert Render/Capture to string
- RoleToString(): Console/Multimedia/Communications info
- DeviceStateToString(): Active/Disabled/Unplugged states

### Memory Management
- PropVariant cleanup
- Proper cleanup for COM objects
- CoTaskMemFree to prevent memory leaks

### Constants and Definitions
- Device state constants (DEVICE_STATE_ACTIVE etc.)
- Property key definitions (PKEY_Device_FriendlyName)
- PropVariant type constants (VT_LPWSTR etc.)

## Usage

```pascal
AudioMonitor := TAudioDeviceMonitor.Create(Self);
AudioMonitor.OnDefaultDeviceChanged := MyDeviceChangeHandler;

procedure TForm1.MyDeviceChangeHandler(Sender: TObject; const DeviceInfo: TAudioDeviceInfo);
begin
  ShowMessage('Audio device changed: ' + DeviceInfo.DeviceName);
end;
```

## Key Improvements Over Original

The enhanced version provides better reliability for detecting audio device changes like switching from speakers to headphones, with proper error handling and thread-safe event notifications.
