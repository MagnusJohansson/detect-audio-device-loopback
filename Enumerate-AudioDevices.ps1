# Get-AudioDevices-Basic.ps1
# PowerShell script to enumerate all audio interfaces using MMDeviceEnumerator
# Compatible with Windows 10/11 - Basic version without property access

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace AudioEnum
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
        int OpenPropertyStore(uint stgmAccess, out IntPtr ppProperties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        int GetState(out uint pdwState);
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumeratorComObject
    {
    }

    public class AudioDeviceInfo
    {
        public string Id;
        public string State;
    }

    public class AudioDeviceEnumerator
    {
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
Write-Host "`nAudio Device Enumerator using MMDeviceEnumerator" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# Display menu
Write-Host "Select enumeration mode:" -ForegroundColor Yellow
Write-Host "1. All devices (Render + Capture)" -ForegroundColor White
Write-Host "2. Render devices only (Playback)" -ForegroundColor White
Write-Host "3. Capture devices only (Recording)" -ForegroundColor White
Write-Host "4. Active devices only" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter your choice (1-4)"

# Set parameters based on choice
$dataFlow = [AudioEnum.EDataFlow]::eAll
$deviceState = [AudioEnum.EDeviceState]::DEVICE_STATEMASK_ALL

switch ($choice) {
    "2" { $dataFlow = [AudioEnum.EDataFlow]::eRender }
    "3" { $dataFlow = [AudioEnum.EDataFlow]::eCapture }
    "4" { $deviceState = [AudioEnum.EDeviceState]::DEVICE_STATE_ACTIVE }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Audio Device Enumeration" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

try {
    # Enumerate devices
    $devices = [AudioEnum.AudioDeviceEnumerator]::EnumerateDevices($dataFlow, $deviceState)
    
    Write-Host "Found $($devices.Count) audio device(s)`n" -ForegroundColor Green
    
    # Display device information
    for ($i = 0; $i -lt $devices.Count; $i++) {
        $device = $devices[$i]
        
        Write-Host "Device #$($i + 1)" -ForegroundColor Yellow
        Write-Host "----------------------------------------" -ForegroundColor Yellow
        Write-Host "Device ID: $($device.Id)" -ForegroundColor White
        Write-Host "State: $($device.State)" -ForegroundColor White
        Write-Host ""
    }
    
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
