# Check-AudioLoopback.ps1
# PowerShell script to enumerate audio devices and check loopback support
# Compatible with Windows 10/11

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace AudioLoopback
{
    public enum EDataFlow
    {
        eRender = 0,
        eCapture = 1,
        eAll = 2
    }

    public enum EDeviceState
    {
        DEVICE_STATE_ACTIVE = 0x00000001,
        DEVICE_STATE_DISABLED = 0x00000002,
        DEVICE_STATE_NOTPRESENT = 0x00000004,
        DEVICE_STATE_UNPLUGGED = 0x00000008,
        DEVICE_STATEMASK_ALL = 0x0000000F
    }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow dataFlow, uint dwStateMask, out IMMDeviceCollection ppDevices);
    }

    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceCollection
    {
        int GetCount(out uint pcDevices);
        int Item(uint nDevice, out IMMDevice ppDevice);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        int Activate(ref Guid iid, uint dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore(uint stgmAccess, out IPropertyStore ppProperties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        int GetState(out uint pdwState);
    }

    [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {
        int GetCount(out uint cProps);
        int GetAt(uint iProp, out PROPERTYKEY pkey);
        int GetValue(ref PROPERTYKEY key, out PropVariant pv);
        int SetValue(ref PROPERTYKEY key, ref PropVariant propvar);
        int Commit();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PropVariant
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pwszVal;
        [FieldOffset(8)] public uint uintVal;
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumeratorComObject
    {
    }

    public class AudioDeviceInfo
    {
        public string Id;
        public string FriendlyName;
        public string State;
        public bool IsRenderDevice;
        public bool SupportsLoopback;
        public string LoopbackInfo;
    }

    public class AudioDeviceEnumerator
    {
        // IAudioClient GUID for activation
        private static readonly Guid IID_IAudioClient = new Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2");

        public static List<AudioDeviceInfo> EnumerateDevices(EDataFlow dataFlow, EDeviceState deviceState)
        {
            List<AudioDeviceInfo> devices = new List<AudioDeviceInfo>();
            IMMDeviceEnumerator deviceEnumerator = null;
            IMMDeviceCollection deviceCollection = null;

            try
            {
                deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
                int hr = deviceEnumerator.EnumAudioEndpoints(dataFlow, (uint)deviceState, out deviceCollection);
                
                if (hr != 0)
                {
                    throw new Exception(string.Format("EnumAudioEndpoints failed with HRESULT: 0x{0:X8}", hr));
                }

                uint count;
                hr = deviceCollection.GetCount(out count);
                
                if (hr != 0)
                {
                    throw new Exception(string.Format("GetCount failed with HRESULT: 0x{0:X8}", hr));
                }

                for (uint i = 0; i < count; i++)
                {
                    IMMDevice device = null;
                    try
                    {
                        hr = deviceCollection.Item(i, out device);
                        if (hr != 0 || device == null) continue;

                        AudioDeviceInfo deviceInfo = new AudioDeviceInfo();

                        // Get Device ID
                        string deviceId;
                        hr = device.GetId(out deviceId);
                        if (hr == 0)
                        {
                            deviceInfo.Id = deviceId;
                            
                            // Determine if it's a render device from the ID
                            // Render devices have {0.0.0.00000000} prefix
                            // Capture devices have {0.0.1.00000000} prefix
                            deviceInfo.IsRenderDevice = deviceId.StartsWith("{0.0.0.");
                        }

                        // Get Device State
                        uint state;
                        hr = device.GetState(out state);
                        if (hr == 0)
                        {
                            List<string> states = new List<string>();
                            if ((state & 0x00000001) != 0) states.Add("ACTIVE");
                            if ((state & 0x00000002) != 0) states.Add("DISABLED");
                            if ((state & 0x00000004) != 0) states.Add("NOTPRESENT");
                            if ((state & 0x00000008) != 0) states.Add("UNPLUGGED");
                            deviceInfo.State = states.Count > 0 ? string.Join(", ", states.ToArray()) : "UNKNOWN";
                        }

                        // Get Friendly Name from property store
                        IPropertyStore propertyStore = null;
                        try
                        {
                            hr = device.OpenPropertyStore(0, out propertyStore);
                            if (hr == 0 && propertyStore != null)
                            {
                                // PKEY_Device_FriendlyName
                                PROPERTYKEY key = new PROPERTYKEY();
                                key.fmtid = new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0);
                                key.pid = 14;
                                
                                PropVariant propVariant;
                                hr = propertyStore.GetValue(ref key, out propVariant);
                                if (hr == 0 && propVariant.vt == 31) // VT_LPWSTR
                                {
                                    deviceInfo.FriendlyName = Marshal.PtrToStringUni(propVariant.pwszVal);
                                }
                            }
                        }
                        catch
                        {
                            // Property store access may fail, continue without friendly name
                        }
                        finally
                        {
                            if (propertyStore != null)
                            {
                                Marshal.ReleaseComObject(propertyStore);
                            }
                        }

                        // Check loopback support
                        // All Windows render (playback) devices support loopback capture via WASAPI
                        if (deviceInfo.IsRenderDevice)
                        {
                            deviceInfo.SupportsLoopback = true;
                            if (deviceInfo.State.Contains("ACTIVE"))
                            {
                                deviceInfo.LoopbackInfo = "Loopback available (WASAPI Loopback Mode)";
                            }
                            else
                            {
                                deviceInfo.LoopbackInfo = "Loopback supported but device not active";
                            }
                        }
                        else
                        {
                            deviceInfo.SupportsLoopback = false;
                            deviceInfo.LoopbackInfo = "Capture device - loopback not applicable";
                        }

                        devices.Add(deviceInfo);
                    }
                    finally
                    {
                        if (device != null)
                        {
                            Marshal.ReleaseComObject(device);
                        }
                    }
                }
            }
            finally
            {
                if (deviceCollection != null)
                {
                    Marshal.ReleaseComObject(deviceCollection);
                }
                if (deviceEnumerator != null)
                {
                    Marshal.ReleaseComObject(deviceEnumerator);
                }
            }

            return devices;
        }
    }
}
"@

# Main Script
Write-Host "`nAudio Device Loopback Checker" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# Display menu
Write-Host "Select enumeration mode:" -ForegroundColor Yellow
Write-Host "1. All devices (Render + Capture)" -ForegroundColor White
Write-Host "2. Render devices only (Playback) - Loopback capable" -ForegroundColor White
Write-Host "3. Capture devices only (Recording)" -ForegroundColor White
Write-Host "4. Active render devices with loopback" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter your choice (1-4)"

# Set parameters based on choice
$dataFlow = [AudioLoopback.EDataFlow]::eAll
$deviceState = [AudioLoopback.EDeviceState]::DEVICE_STATEMASK_ALL

switch ($choice) {
    "2" { $dataFlow = [AudioLoopback.EDataFlow]::eRender }
    "3" { $dataFlow = [AudioLoopback.EDataFlow]::eCapture }
    "4" { 
        $dataFlow = [AudioLoopback.EDataFlow]::eRender
        $deviceState = [AudioLoopback.EDeviceState]::DEVICE_STATE_ACTIVE
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Audio Device Enumeration" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

try {
    # Enumerate devices
    $devices = [AudioLoopback.AudioDeviceEnumerator]::EnumerateDevices($dataFlow, $deviceState)
    
    Write-Host "Found $($devices.Count) audio device(s)`n" -ForegroundColor Green
    
    $loopbackCount = 0
    $activeLoopbackCount = 0
    
    # Display device information
    for ($i = 0; $i -lt $devices.Count; $i++) {
        $device = $devices[$i]
        
        # Color code based on loopback support
        $headerColor = "Yellow"
        if ($device.SupportsLoopback -and $device.State.Contains("ACTIVE")) {
            $headerColor = "Green"
            $activeLoopbackCount++
        }
        
        if ($device.SupportsLoopback) {
            $loopbackCount++
        }
        
        Write-Host "Device #$($i + 1)" -ForegroundColor $headerColor
        Write-Host "----------------------------------------" -ForegroundColor $headerColor
        
        if ($device.FriendlyName) {
            Write-Host "Name: $($device.FriendlyName)" -ForegroundColor White
        }
        
        Write-Host "Device ID: $($device.Id)" -ForegroundColor Gray
        Write-Host "State: $($device.State)" -ForegroundColor White
        Write-Host "Type: $(if ($device.IsRenderDevice) { 'Render (Playback)' } else { 'Capture (Recording)' })" -ForegroundColor White
        
        # Highlight loopback info
        if ($device.SupportsLoopback) {
            Write-Host "Loopback: $($device.LoopbackInfo)" -ForegroundColor Cyan
        } else {
            Write-Host "Loopback: $($device.LoopbackInfo)" -ForegroundColor DarkGray
        }
        
        Write-Host ""
    }
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "  Total devices: $($devices.Count)" -ForegroundColor White
    Write-Host "  Loopback capable: $loopbackCount" -ForegroundColor White
    Write-Host "  Active with loopback: $activeLoopbackCount" -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Cyan
}
catch {
    Write-Error "An error occurred: $_"
    Write-Host "`nError details:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
}
