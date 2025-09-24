using System;
using System.Runtime.InteropServices;

/// <summary>
/// VolumeControl Application
/// 
/// This application provides precise control over Windows system volume using Windows API calls.
/// It supports absolute volume levels (0-50), mute/unmute functionality with volume restoration,
/// and quiet operation mode. Volume levels are persisted in the Windows registry for restoration.
/// 
/// Key Features:
/// - Absolute volume control (0-50 scale)
/// - Mute with automatic volume level storage
/// - Volume restoration via /default switch
/// - Quiet mode operation
/// - Registry-based persistence (single setting)
/// - Designed to work with scheduled task lockscreen on/off
/// - Windows Application mode - no console window unless run from command line
/// 
/// Author: Simon Bond
/// Date: 24th September 2025
/// Version 1.0
/// </summary>
public class VolumeControl
{
    #region Windows API Declarations
    
    /// <summary>
    /// Gets the handle of the currently active (foreground) window
    /// </summary>
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    /// <summary>
    /// Retrieves the process ID of the thread that created the specified window
    /// </summary>
    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    /// <summary>
    /// Gets the handle of the console window associated with the current process
    /// </summary>
    [DllImport("kernel32.dll")]
    static extern IntPtr GetConsoleWindow();

    /// <summary>
    /// Allocates a new console for the current process
    /// </summary>
    [DllImport("kernel32.dll")]
    static extern bool AllocConsole();

    /// <summary>
    /// Attaches the process to the console of the specified process
    /// </summary>
    [DllImport("kernel32.dll")]
    static extern bool AttachConsole(int dwProcessId);

    /// <summary>
    /// Frees the console associated with the calling process
    /// </summary>
    [DllImport("kernel32.dll")]
    static extern bool FreeConsole();

    /// <summary>
    /// Shows or hides a window (used to control console window visibility)
    /// </summary>
    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    // Window show state constants
    const int SW_HIDE = 0;  // Hide the window
    const int SW_SHOW = 5;  // Show the window normally

    // Special process ID to attach to parent process console
    const int ATTACH_PARENT_PROCESS = -1;

    /// <summary>
    /// Sends a message to a window or windows. This is the core function for volume control.
    /// We use this to send application commands (like volume up/down) to the Windows shell.
    /// 
    /// Why SendMessage is superior to SendKeys:
    /// 1. Direct System Integration: SendMessage with WM_APPCOMMAND sends commands directly 
    ///    to the Windows message system, bypassing keyboard simulation entirely.
    /// 2. No Focus Dependencies: Works regardless of which window has focus, unlike SendKeys 
    ///    which requires the target application to be active.
    /// 3. Immune to Input Blocking: Cannot be intercepted or blocked by security software, 
    ///    input hooks, or other applications.
    /// 4. Precise Control: Sends exact system commands rather than simulating keystrokes 
    ///    that might be interpreted differently.
    /// 5. No Timing Issues: Executes immediately without timing dependencies or the need 
    ///    to wait for keystrokes to be processed.
    /// </summary>
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    /// <summary>
    /// Finds a window by its class name and/or window name.
    /// We use this to find the Windows shell (taskbar) for reliable message delivery.
    /// </summary>
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    #endregion

    #region Registry API Declarations
    
    /// <summary>
    /// Registry P/Invoke declarations for storing and retrieving volume settings.
    /// We use direct registry API calls instead of Microsoft.Win32.Registry to avoid
    /// external dependencies in .NET 8.
    /// </summary>

    /// <summary>
    /// Creates a new registry key or opens an existing one
    /// </summary>
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    private static extern int RegCreateKeyEx(UIntPtr hKey, string lpSubKey, uint Reserved, string lpClass,
        uint dwOptions, uint samDesired, IntPtr lpSecurityAttributes, out UIntPtr phkResult, out uint lpdwDisposition);

    /// <summary>
    /// Opens an existing registry key for reading
    /// </summary>
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    private static extern int RegOpenKeyEx(UIntPtr hKey, string lpSubKey, uint ulOptions, uint samDesired, out UIntPtr phkResult);

    /// <summary>
    /// Sets a value in an open registry key
    /// </summary>
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    private static extern int RegSetValueEx(UIntPtr hKey, string lpValueName, uint Reserved, uint dwType, byte[] lpData, uint cbData);

    /// <summary>
    /// Retrieves a value from an open registry key
    /// </summary>
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    private static extern int RegQueryValueEx(UIntPtr hKey, string lpValueName, IntPtr lpReserved, out uint lpType, byte[] lpData, ref uint lpcbData);

    /// <summary>
    /// Closes an open registry key handle
    /// </summary>
    [DllImport("advapi32.dll")]
    private static extern int RegCloseKey(UIntPtr hKey);

    #endregion

    #region Constants and Configuration

    // Windows Message Constants for volume control
    private const uint WM_APPCOMMAND = 0x319;           // Application command message
    private const uint APPCOMMAND_VOLUME_UP = 0x0a;     // Volume up command
    private const uint APPCOMMAND_VOLUME_DOWN = 0x09;   // Volume down command  
    private const uint APPCOMMAND_VOLUME_MUTE = 0x08;   // Volume mute command

    // Application settings
    private static bool suppressOutput = false;          // Controls console output visibility
    private static bool hasConsole = false;              // Tracks if we have console access
    private const int MAX_VOLUME_LEVEL = 50;            // Maximum volume level (0-50 scale)

    // Registry constants for storing volume settings
    private static readonly UIntPtr HKEY_CURRENT_USER = new UIntPtr(0x80000001u);  // Registry root
    private const uint KEY_SET_VALUE = 0x0002;          // Permission to write registry values
    private const uint KEY_QUERY_VALUE = 0x0001;       // Permission to read registry values
    private const uint REG_DWORD = 4;                  // Registry data type for 32-bit integers
    private const uint REG_OPTION_NON_VOLATILE = 0;    // Registry key persists after reboot

    // Registry paths and value names - SIMPLIFIED TO ONE SETTING
    private const string REGISTRY_KEY_PATH = @"SOFTWARE\Morgan Stanley\SpeakerVolume";
    private const string VOLUME_LEVEL_VALUE = "VolumeLevel";  // Single registry value for volume storage

    #endregion

    #region Console Management

    /// <summary>
    /// Initializes console output. This method handles different scenarios:
    /// 1. If launched from command prompt - attaches to parent console
    /// 2. If launched standalone with --debug - creates new console
    /// 3. If launched standalone normally - no console (silent operation)
    /// </summary>
    /// <param name="forceConsole">Force creation of console even when launched standalone</param>
    private static void InitializeConsole(bool forceConsole = false)
    {
        // First, try to attach to parent process console (command prompt)
        if (AttachConsole(ATTACH_PARENT_PROCESS))
        {
            hasConsole = true;
            // Successfully attached to parent console (launched from command prompt)
            return;
        }

        // If we couldn't attach to parent and force console is requested
        if (forceConsole)
        {
            if (AllocConsole())
            {
                hasConsole = true;
                // Redirect standard streams to the new console
                Console.SetOut(new System.IO.StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
                Console.SetError(new System.IO.StreamWriter(Console.OpenStandardError()) { AutoFlush = true });
            }
        }

        // If neither attach nor alloc worked, we'll run silently (hasConsole remains false)
    }

    /// <summary>
    /// Cleans up console resources when the application exits
    /// </summary>
    private static void CleanupConsole()
    {
        if (hasConsole)
        {
            FreeConsole();
        }
    }

    #endregion

    #region Main Application Entry Point

    /// <summary>
    /// Main application entry point. Handles command line argument parsing and determines
    /// the operation mode (set volume, restore default, or show usage).
    /// 
    /// Supported command line arguments:
    /// - Numeric value (0-50): Sets absolute volume level
    /// - /default: Restores previously stored volume level
    /// - --quiet or -q: Suppresses console output
    /// - --debug: Forces console creation for debugging (when run standalone)
    /// </summary>
    static void Main(string[] args)
    {
        // Initialize command line parsing variables
        int volumeLevel = -1;        // Target volume level (-1 means not specified)
        bool useDefault = false;     // Whether to restore default volume
        bool forceConsole = false;   // Whether to force console creation

        // Parse all command line arguments
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--quiet" || args[i] == "-q")
            {
                // Enable quiet mode - suppresses all console output
                suppressOutput = true;
            }
            else if (args[i] == "/default")
            {
                // Enable default restoration mode
                useDefault = true;
            }
            else if (args[i] == "--debug")
            {
                // Force console creation for debugging
                forceConsole = true;
            }
            else if (int.TryParse(args[i], out int vol))
            {
                // Parse numeric volume level
                volumeLevel = vol;
            }
        }

        // Initialize console based on how we were launched
        InitializeConsole(forceConsole);

        // Handle console window visibility based on quiet mode
        if (suppressOutput && hasConsole)
        {
            IntPtr consoleWindow = GetConsoleWindow();
            if (consoleWindow != IntPtr.Zero)
            {
                ShowWindow(consoleWindow, SW_HIDE);
            }
        }

        try
        {
            // Determine operation mode and execute appropriate action
            if (useDefault)
            {
                // Restore volume from registry
                int storedVolume = GetStoredVolumeLevel();
                if (storedVolume >= 0) // Only restore if we have a valid volume
                {
                    WriteOutput($"Restoring volume to stored level: {storedVolume}");
                    SetVolume(storedVolume);
                }
                else
                {
                    WriteOutput("No valid stored volume level found in registry.");
                }
            }
            else if (volumeLevel != -1)
            {
                // Set specific volume level
                SetVolume(volumeLevel);
            }
            else
            {
                // No valid arguments provided - show usage information
                ShowUsage();
            }
        }
        finally
        {
            // Clean up console resources
            CleanupConsole();
        }
    }

    /// <summary>
    /// Displays usage information and command line syntax
    /// </summary>
    private static void ShowUsage()
    {
        WriteOutput("Usage: VolumeControl <volume_level> [--quiet|-q] [--debug]");
        WriteOutput("       VolumeControl /default [--quiet|-q] [--debug]");
        WriteOutput($"volume_level: An integer between 0 and {MAX_VOLUME_LEVEL} (absolute volume level).");
        WriteOutput($"  0 = Muted, 25 = Half volume, {MAX_VOLUME_LEVEL} = Maximum volume");
        WriteOutput("/default: Restore volume to previously stored level.");
        WriteOutput("--quiet or -q: Suppress console output.");
        WriteOutput("--debug: Force console window creation (for debugging when run standalone).");
    }

    #endregion

    #region Volume Control Logic

    /// <summary>
    /// Core volume control method. Sets the system volume to the specified level using
    /// a reset-and-build approach for maximum accuracy.
    /// 
    /// The algorithm works as follows:
    /// 1. If muting (volume 0), first store the current actual system volume level
    /// 2. Reset volume to 0 by sending multiple volume-down commands
    /// 3. If target is 0, stop here (muted)
    /// 4. If target > 0, send volume-up commands to reach the target level
    /// 
    /// Registry Usage:
    /// - When setting volume > 0: Store the volume level in registry
    /// - When setting volume = 0: Store the CURRENT actual system volume first, then mute
    /// - When using /default: Restore the volume level from registry
    /// </summary>
    /// <param name="volumeLevel">Target volume level (0-50)</param>
    public static void SetVolume(int volumeLevel)
    {
        // Validate input range
        if (volumeLevel < 0 || volumeLevel > MAX_VOLUME_LEVEL)
        {
            WriteOutput($"Volume level must be between 0 and {MAX_VOLUME_LEVEL}.");
            return;
        }

        try
        {
            WriteOutput($"Setting volume to level {volumeLevel} (out of {MAX_VOLUME_LEVEL})");

            // Special handling for mute operation (volume level 0)
            // Store the CURRENT actual system volume before muting
            if (volumeLevel == 0)
            {
                // Get the current actual system volume
                float currentSystemVolume = GetMasterVolume();
                int currentVolumeLevel = (int)(currentSystemVolume * MAX_VOLUME_LEVEL); // Convert to 0-50 scale
                
                WriteOutput($"Storing current system volume level ({currentVolumeLevel}) before muting.");
                StoreVolumeLevel(currentVolumeLevel);
            }

            // Get window handle for sending volume commands
            // We prioritize the Windows shell (taskbar) as it reliably handles volume commands
            IntPtr hWnd = FindWindow("Shell_TrayWnd", null);
            if (hWnd == IntPtr.Zero)
            {
                // Fallback to foreground window if shell not found
                hWnd = GetForegroundWindow();
            }

            // PHASE 1: Reset to baseline (volume 0)
            // This ensures we start from a known state regardless of current volume
            WriteOutput("Resetting volume to level 0 (mute)...");
            for (int i = 0; i < 50; i++) // Send 50 volume-down commands to ensure we hit 0
            {
                SendMessage(hWnd, WM_APPCOMMAND, IntPtr.Zero, (IntPtr)(APPCOMMAND_VOLUME_DOWN << 16));
                System.Threading.Thread.Sleep(10); // Small delay to ensure command processing
            }

            // PHASE 2: Handle mute case (target volume is 0)
            if (volumeLevel == 0)
            {
                WriteOutput("Volume set to level 0 (muted)");
                return; // Don't store 0 in registry - we already stored the previous volume
            }

            // PHASE 3: Build up to target volume
            // Each volume-up command increases volume by approximately 2% of system maximum
            // Since our scale is 0-50, each step corresponds to one volume-up command
            int stepsNeeded = volumeLevel;

            WriteOutput($"Increasing volume to level {volumeLevel} ({stepsNeeded} steps)...");

            for (int i = 0; i < stepsNeeded; i++)
            {
                SendMessage(hWnd, WM_APPCOMMAND, IntPtr.Zero, (IntPtr)(APPCOMMAND_VOLUME_UP << 16));
                System.Threading.Thread.Sleep(20); // Slightly longer delay for volume-up commands
            }

            WriteOutput($"Volume set to level {volumeLevel} (out of {MAX_VOLUME_LEVEL})");

            // PHASE 4: Store the new volume level (only for non-zero volumes)
            StoreVolumeLevel(volumeLevel);
        }
        catch (Exception ex)
        {
            WriteOutput($"Error setting volume: {ex.Message}");
        }
    }

    #endregion

    #region Registry Operations - SIMPLIFIED TO ONE SETTING

    /// <summary>
    /// Stores a volume level in the single registry entry.
    /// This is used for both current volume storage and restoration.
    /// </summary>
    /// <param name="volumeLevel">Volume level to store (1-50)</param>
    private static void StoreVolumeLevel(int volumeLevel)
    {
        try
        {
            UIntPtr hKey;
            uint disposition;
            
            // Create or open the registry key
            int result = RegCreateKeyEx(HKEY_CURRENT_USER, REGISTRY_KEY_PATH, 0, null, 
                REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, IntPtr.Zero, out hKey, out disposition);
            
            if (result == 0) // Success
            {
                // Convert integer to byte array for storage
                byte[] data = BitConverter.GetBytes(volumeLevel);
                
                // Store the value
                result = RegSetValueEx(hKey, VOLUME_LEVEL_VALUE, 0, REG_DWORD, data, (uint)data.Length);
                
                // Always close the registry key handle
                RegCloseKey(hKey);
                
                if (result == 0)
                {
                    WriteOutput($"Volume level {volumeLevel} stored in registry.");
                }
                else
                {
                    WriteOutput($"Error storing volume level in registry. Error code: {result}");
                }
            }
            else
            {
                WriteOutput($"Error creating registry key. Error code: {result}");
            }
        }
        catch (Exception ex)
        {
            WriteOutput($"Error storing volume level in registry: {ex.Message}");
        }
    }

    /// <summary>
    /// Retrieves the stored volume level from the single registry entry.
    /// This is used by the /default switch to restore volume.
    /// </summary>
    /// <returns>Volume level (1-50) or -1 if not found</returns>
    private static int GetStoredVolumeLevel()
    {
        try
        {
            UIntPtr hKey;
            
            // Open the registry key for reading
            int result = RegOpenKeyEx(HKEY_CURRENT_USER, REGISTRY_KEY_PATH, 0, KEY_QUERY_VALUE, out hKey);
            
            if (result == 0) // Success
            {
                uint type;
                byte[] data = new byte[4]; // DWORD is 4 bytes
                uint dataSize = (uint)data.Length;
                
                // Read the value
                result = RegQueryValueEx(hKey, VOLUME_LEVEL_VALUE, IntPtr.Zero, out type, data, ref dataSize);
                
                // Always close the registry key handle
                RegCloseKey(hKey);
                
                // Verify successful read and correct data type
                if (result == 0 && type == REG_DWORD)
                {
                    // Convert byte array back to integer
                    return BitConverter.ToInt32(data, 0);
                }
            }
        }
        catch (Exception ex)
        {
            WriteOutput($"Error reading volume level from registry: {ex.Message}");
        }
        
        return -1; // Return -1 to indicate value not found or error occurred
    }

    #endregion

    #region Windows Core Audio API for Getting Current Volume

    internal static class ComCLSIDs
    {
        public const string MMDeviceEnumerator = "BCDE0395-E52F-467C-8E3D-C4579291692E";
    }

    internal static class ComIIDs
    {
        public const string IMMDeviceEnumerator = "A95664D2-9614-4F35-A746-DE8DB63617E6";
        public const string IAudioEndpointVolume = "5CDF2C82-841E-4546-9722-0CF74078229A";
    }

    // COM Interfaces
    [ComImport]
    [Guid(ComIIDs.IMMDeviceEnumerator)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        [PreserveSig]
        int EnumAudioEndpoints(int dataFlow, int dwStateMask, out IntPtr ppDevices);

        [PreserveSig]
        int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice? ppEndpoint);

        [PreserveSig]
        int GetDevice(string pwstrId, out IMMDevice? ppDevice);

        [PreserveSig]
        int RegisterEndpointNotificationCallback(IntPtr pClient);

        [PreserveSig]
        int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        [PreserveSig]
        int Activate(ref Guid iid, uint dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        [PreserveSig]
        int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);

        [PreserveSig]
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

        [PreserveSig]
        int GetState(out int pdwState);
    }

    [ComImport]
    [Guid(ComIIDs.IAudioEndpointVolume)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolume
    {
        [PreserveSig]
        int RegisterControlChangeNotify(IntPtr pNotify);

        [PreserveSig]
        int UnregisterControlChangeNotify(IntPtr pNotify);

        [PreserveSig]
        int GetChannelCount(out int pnChannelCount);

        [PreserveSig]
        int SetMasterVolumeLevel(float fLevelDB, ref Guid pguidEventContext);

        [PreserveSig]
        int SetMasterVolumeLevelScalar(float fLevel, ref Guid pguidEventContext);

        [PreserveSig]
        int GetMasterVolumeLevel(out float pfLevelDB);

        [PreserveSig]
        int GetMasterVolumeLevelScalar(out float pfLevel);

        [PreserveSig]
        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, ref Guid pguidEventContext);

        [PreserveSig]
        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, ref Guid pguidEventContext);

        [PreserveSig]
        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);

        [PreserveSig]
        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);

        [PreserveSig]
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid pguidEventContext);

        [PreserveSig]
        int GetMute(out bool pbMute);

        [PreserveSig]
        int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);

        [PreserveSig]
        int VolumeStepUp(ref Guid pguidEventContext);

        [PreserveSig]
        int VolumeStepDown(ref Guid pguidEventContext);

        [PreserveSig]
        int QueryHardwareSupport(out uint pdwHardwareSupportMask);

        [PreserveSig]
        int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }

    /// <summary>
    /// Gets the current system master volume level using Windows Core Audio API.
    /// This provides the actual current system volume, not an estimated value.
    /// </summary>
    /// <returns>Volume level as a float between 0.0 and 1.0</returns>
    static float GetMasterVolume()
    {
        IMMDeviceEnumerator? deviceEnumerator = null;
        IMMDevice? defaultDevice = null;
        IAudioEndpointVolume? endpointVolume = null;
        try
        {
            // Instantiate the MMDeviceEnumerator COM object
            deviceEnumerator = (IMMDeviceEnumerator)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid(ComCLSIDs.MMDeviceEnumerator))!)!;

            // Get the default audio rendering device
            const int eRender = 0; // DataFlow.Render
            const int eMultimedia = 1; // Role.Multimedia
            deviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, out defaultDevice);

            if (defaultDevice == null)
            {
                throw new InvalidOperationException("Default audio device not found.");
            }

            // Activate the IAudioEndpointVolume interface
            Guid iid = new Guid(ComIIDs.IAudioEndpointVolume);
            const uint CLSCTX_INPROC_SERVER = 1;
            defaultDevice.Activate(ref iid, CLSCTX_INPROC_SERVER, IntPtr.Zero, out object o);
            endpointVolume = (IAudioEndpointVolume)o;

            // Get the master volume level scalar (0.0 to 1.0)
            endpointVolume.GetMasterVolumeLevelScalar(out float volumeLevel);
            return volumeLevel;
        }
        finally
        {
            // Clean up COM objects
            if (endpointVolume != null) Marshal.ReleaseComObject(endpointVolume);
            if (defaultDevice != null) Marshal.ReleaseComObject(defaultDevice);
            if (deviceEnumerator != null) Marshal.ReleaseComObject(deviceEnumerator);
        }
    }

    #endregion

    #region Utility Methods

    /// <summary>
    /// Conditional output method that respects the quiet mode setting and console availability.
    /// All console output in the application goes through this method.
    /// </summary>
    /// <param name="message">Message to write to console</param>
    private static void WriteOutput(string message)
    {
        // Only write output if we're not in quiet mode AND we have console access
        if (!suppressOutput && hasConsole)
        {
            Console.WriteLine(message);
        }
    }

    #endregion
}