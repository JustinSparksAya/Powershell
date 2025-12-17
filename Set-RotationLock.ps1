param(
  [switch]$on,
  [switch]$off
)

$Source = @"
using System;
using System.Runtime.InteropServices;

public class SystemRotation {
    // Import the function to check rotation state
    // pState: 0 = Auto-rotation disabled (Locked), 1 = Auto-rotation enabled
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetAutoRotationState(out int pState);

    // Your existing SetAutoRotation method
    [DllImport("user32.dll", EntryPoint = "#2507", SetLastError = true)]
    public static extern bool SetAutoRotation(bool bEnable);

    public static string GetCurrentStatus() {
        int state;
        if (GetAutoRotationState(out state)) {
            return (state == 1) ? "Auto-Rotation Enabled (Lock is OFF)" : "Auto-Rotation Disabled (Lock is ON)";
        }
        return "Unable to retrieve rotation state. The device may not support it.";
    }
}
"@

Add-Type -TypeDefinition $Source

# Get and display the current status
Write-Host "Current Rotation Status: $([SystemRotation]::GetCurrentStatus())"
if($on){
  [SystemRotation]::SetAutoRotation($false) # Disable auto-rotation (Lock ON)
}elseif($off){
  [SystemRotation]::SetAutoRotation($true) # Disable auto-rotation (Lock OFF)
}
Write-Host "New Rotation Status: $([SystemRotation]::GetCurrentStatus())"
