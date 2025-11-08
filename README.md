# detect-audio-device-loopback

PowerShell script that enumerates all audio devices on Windows 10/11 using MMDeviceEnumerator.

## Features

- Enumerates all audio interfaces using Windows Core Audio API (MMDeviceEnumerator)
- Displays device information including:
  - Device ID (unique identifier)
  - Device State (Active, Disabled, Not Present, Unplugged)
- Filter by device type (Render/Playback, Capture/Recording, or All)
- Filter by device state (All states or Active only)
- Pure PowerShell implementation with C# COM interop

## Usage

Run the script in PowerShell:

```powershell
.\Enumerate-AudioDevices.ps1
```

The script will present a menu with the following options:
1. All devices (Render + Capture)
2. Render devices only (Playback)
3. Capture devices only (Recording)
4. Active devices only

### Example Output

```
Found 20 audio device(s)

Device #1
----------------------------------------
Device ID: {0.0.0.00000000}.{b65a1f33-a4b7-4151-a207-50294234d8b6}
State: ACTIVE

Device #2
----------------------------------------
Device ID: {0.0.1.00000000}.{6b320a69-4807-4ba4-9c61-2a55ac5b80b2}
State: ACTIVE
```

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- No administrator privileges required

## Technical Details

The script uses the Windows Core Audio API through COM interop:
- **MMDeviceEnumerator**: COM object (CLSID: BCDE0395-E52F-467C-8E3D-C4579291692E) for enumerating audio endpoint devices
- **IMMDeviceEnumerator**: Interface for device enumeration
- **IMMDeviceCollection**: Interface for accessing the collection of devices
- **IMMDevice**: Interface representing an audio endpoint device
- Supports all device states (Active, Disabled, Not Present, Unplugged)
- Supports all data flow directions (Render, Capture, All)

## Loopback Support Detection

To check which devices support loopback recording, use the enhanced script:

```powershell
.\Check-AudioLoopback.ps1
```

### How Loopback Works in Windows

**All Windows render (playback) devices support loopback capture** through WASAPI Loopback Mode. This allows you to record the audio output from any application.

**Key Points:**
- **Render devices** (speakers, headphones) support loopback - identified by `{0.0.0.` prefix in Device ID
- **Capture devices** (microphones) do not support loopback - identified by `{0.0.1.` prefix in Device ID
- Loopback is only available when the device is in **ACTIVE** state
- Loopback captures the audio stream before it reaches the hardware (system audio)

### Example Output

```
Device #13
----------------------------------------
Name: Speakers (Qualcomm(R) Aqstic(TM) Audio Adapter Device)
Device ID: {0.0.0.00000000}.{b65a1f33-a4b7-4151-a207-50294234d8b6}
State: ACTIVE
Type: Render (Playback)
Loopback: Loopback available (WASAPI Loopback Mode)
```

## Notes

- Basic enumeration script (`Enumerate-AudioDevices.ps1`) shows all devices with IDs and states
- Loopback checker script (`Check-AudioLoopback.ps1`) identifies which devices support loopback recording
- Device IDs can be used with WASAPI APIs to capture loopback audio programmatically
