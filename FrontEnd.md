
Okay, I can help with that! First, I need to understand the project structure to locate the front-end components. I'll start by listing the contents of the project's root directory.


Ran tool
Based on the directory listing, the `Xaml/` directory seems like the most probable location for front-end code, likely using XAML for a WPF or similar Windows UI. The `Start-WithApp.cmd` script might also provide clues on how the UI is launched.

I'll now look into the `Xaml/` directory.


Ran tool
The `Xaml/` directory contains numerous `.xaml` files, which confirms that the front-end is built using XAML. This is typically used with WPF (Windows Presentation Foundation) or UWP (Universal Windows Platform) for creating desktop applications in the Windows ecosystem.

`MainWindow.xaml` is likely the main window of the application. To understand how the UI is structured and how it interacts with the backend logic (which is likely PowerShell in this project), I will now examine `MainWindow.xaml`.


Read file: Xaml/MainWindow.xaml
The `MainWindow.xaml` file defines the main layout of the application. Here's a breakdown:

*   **Window Definition:** It sets up the main window with the title "Cloud API PowerShell Management".
*   **Resources:** It includes a `ResourceDictionary` which likely points to `Themes\Default.xaml`. This means the application's look and feel (styles, colors, etc.) are defined in a separate theme file.
*   **Main Grid Layout:** The window is structured using a `Grid` with rows and columns to arrange different UI elements.
    *   **Menu Bar:** There's a `Menu` (`mnuMain`) with "File" and "Views" options.
        *   The "File" menu includes "Settings", "Tenant Settings", "Release Notes", "About", and "Exit".
    *   **Environment Label:** A `Label` (`lblEnvironment`) to display environment information, initially collapsed.
    *   **View Panel:** A `Grid` (`grdViewPanel`) seems to be the main content area where different views or pages of the application will be displayed.
    *   **Menu Items List:** A `ListBox` (`lstMenuItems`) on the left side, likely used for navigation. It displays items with icons and titles, and their appearance (e.g., text color) can change based on `AccessType` (None, Limited).
    *   **Popups:** There are several `Grid` elements (`grdPopup`, `grdModal`, `grdStatus`) that are initially collapsed. These are likely used for displaying pop-up dialogs, modal forms, or status messages.

**How it likely works with PowerShell:**

XAML defines the UI structure and appearance. The actual logic, data retrieval, and actions performed when you interact with the UI (e.g., click a button, select a menu item) are typically handled by code-behind files (often C# for WPF) or, in this project's case, most likely PowerShell scripts.

The PowerShell scripts would:

1.  **Load the XAML:** Read the `.xaml` files to create the UI elements.
2.  **Populate Data:** Fetch data (e.g., from Intune, Azure AD) and populate the UI elements like list boxes, tables, etc.
3.  **Handle Events:** Attach event handlers to UI elements. For example, when a button is clicked in the XAML UI, a corresponding PowerShell function is executed.
4.  **Update UI:** Modify the UI based on user actions or data changes.

The `Start-WithApp.cmd` file probably initiates this process by running a PowerShell script that loads `MainWindow.xaml` and starts the application.

To give you a more complete picture, I should also look at `Start-WithApp.cmd` and potentially one of the core PowerShell scripts that might be responsible for loading the XAML and managing the UI logic. The `CloudAPIPowerShellManagement.psm1` or `Core.psm1` files seem like good candidates.

Let's start with `Start-WithApp.cmd`.


Read file: Start-WithApp.cmd
The `Start-WithApp.cmd` script executes `Start-IntuneManagement.ps1` using PowerShell 5. It passes parameters like `tenantId`, `appid`, and `secret`, which are placeholders here and would need to be filled in with actual values.

Now, let's examine `Start-IntuneManagement.ps1` to see how it initializes the application.


Read file: Start-IntuneManagement.ps1
The `Start-IntuneManagement.ps1` script does the following:

1.  **Defines Parameters:** It accepts several parameters, including those passed by `Start-WithApp.cmd` (`TenantId`, `AppId`, `Secret`), and others to control its behavior (e.g., `ShowConsoleWindow`, `JSonSettings`).
2.  **Imports Module:** It imports the `CloudAPIPowerShellManagement.psd1` module, which likely contains the core functionality. This psd1 file will point to `CloudAPIPowerShellManagement.psm1`.
3.  **Initializes Management:** It then calls `Initialize-CloudAPIManagement` (a function presumably from the imported module) with the parameter `-View "IntuneGraphAPI"` and passes along all the script's bound parameters (`@param`).

This confirms that the PowerShell module (`CloudAPIPowerShellManagement.psm1` and likely `Core.psm1` which is often imported by the main module) is responsible for setting up the UI. The `-View "IntuneGraphAPI"` parameter suggests that the application starts with a specific view related to Intune Graph API.

So, the front-end is a XAML-based UI, likely WPF, which is dynamically loaded and managed by PowerShell scripts. The `MainWindow.xaml` defines the main application frame, and different "views" (other XAML files) are loaded into the main content area (`grdViewPanel`). PowerShell handles the logic, data, and event handling.

To provide a deeper understanding of how PowerShell specifically interacts with the XAML (e.g., loading XAML, attaching event handlers), I should now look into `CloudAPIPowerShellManagement.psm1`.


Read file: CloudAPIPowerShellManagement.psm1
The `CloudAPIPowerShellManagement.psm1` script, specifically the `Initialize-CloudAPIManagement` function, orchestrates the application startup and UI initialization:

1.  **Console Management:** It includes functions (`Show-Console`, `Hide-Console`) to control the visibility of the PowerShell console window. It typically hides the console when the GUI is running.
2.  **Unblock Files:** An `Unblock-AllFiles` function is called to ensure that downloaded files (like scripts and XAML) are not blocked by Windows security, which can prevent them from running.
3.  **Assembly Loading:** It loads necessary .NET assemblies for WPF (`PresentationFramework`) and Windows Forms (`System.Windows.Forms`, likely for `Application.DoEvents` or other interop).
4.  **Splash Screen:**
    *   If not running in silent mode (`$global:hideUI -ne $true`), it loads and displays `SplashScreen.xaml`.
    *   It uses `Get-Content` to read the XAML file into an `[xml]` object.
    *   `[Windows.Markup.XamlReader]::Load(...)` is the key method used to parse the XAML string and create the actual WPF window object.
    *   It then finds elements within the splash screen (like `txtSplashTitle`, `txtSplashText`) by name using `$global:SplashScreen.FindName("...")` to update their content.
    *   `$global:SplashScreen.Show()` displays the window.
    *   `[System.Windows.Forms.Application]::DoEvents()` is used to keep the UI responsive while initialization tasks are performed in the background.
5.  **Parameter Handling:** It processes parameters like `TenantId`, `AppId`, `Secret`, `ShowConsoleWindow`, and `JSonSettings`, storing them in global variables for later use.
6.  **Core Module Import:** It imports `Core.psm1` using `Import-Module ... -Global`. This `Core.psm1` module likely contains the bulk of the application logic, including how `MainWindow.xaml` is loaded and how different views are managed.
7.  **Start Core App:** Finally, it calls `Start-CoreApp $View`. This function, which must be defined in `Core.psm1`, is responsible for loading the main application window and the initial view specified (e.g., "IntuneGraphAPI").

**In summary, how the front-end works:**

1.  **Launch:** `Start-WithApp.cmd` calls `Start-IntuneManagement.ps1`.
2.  **Initialization Script (`Start-IntuneManagement.ps1`):**
    *   Imports `CloudAPIPowerShellManagement.psd1` (which loads `CloudAPIPowerShellManagement.psm1`).
    *   Calls `Initialize-CloudAPIManagement`.
3.  **Main Module (`CloudAPIPowerShellManagement.psm1` - `Initialize-CloudAPIManagement` function):**
    *   Sets up the environment (hides console, loads .NET assemblies).
    *   Loads and shows `SplashScreen.xaml` by reading the file content and using `[Windows.Markup.XamlReader]::Load()`.
    *   Updates text elements on the splash screen using `FindName()`.
    *   Imports `Core.psm1`.
    *   Calls `Start-CoreApp` (from `Core.psm1`) to load the main application UI.
4.  **Core Logic (`Core.psm1` - assumed):**
    *   The `Start-CoreApp` function will similarly load `MainWindow.xaml` using `XamlReader`.
    *   It will then populate the `lstMenuItems` and handle navigation by loading other `.xaml` files (views like `EndpointManagerPanel.xaml`, `ImportForm.xaml`, etc.) into the `grdViewPanel` of `MainWindow.xaml`.
    *   Event handlers for UI elements (buttons, menu items, list selections) will be PowerShell functions or scriptblocks that are connected to the XAML elements, often by naming convention or by explicitly adding event handlers in PowerShell after loading the XAML.

The front-end is essentially a WPF application where the UI is defined in XAML files, and PowerShell scripts dynamically load these XAML files, interact with the UI elements, and execute the application's logic. This approach allows for a rich graphical user interface powered by PowerShell scripting capabilities for management tasks.

Do you have any more specific questions about a particular part of the front-end or its interaction with PowerShell? For example, we could look into `Core.psm1` to see `Start-CoreApp` or how event handlers might be attached.




Based on the evidence so far:

*   **Primarily WPF:** The extensive use of `.xaml` files, the loading mechanism `[Windows.Markup.XamlReader]::Load()`, and the `PresentationFramework` assembly are all hallmarks of Windows Presentation Foundation (WPF). The structure of the XAML files (e.g., `<Window>`, `<Grid>`, `<Menu>`, `<ListBox>`) is standard WPF.
*   **Windows Forms for specific tasks:** The line `[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")` and the use of `[System.Windows.Forms.Application]::DoEvents()` in `CloudAPIPowerShellManagement.psm1` indicate that some Windows Forms components or functionalities are being utilized.
    *   `Application.DoEvents()` is a common (though sometimes debated) practice to process the message queue and keep the UI responsive during long-running operations in the UI thread, especially when integrating non-WPF blocking calls or in single-threaded apartment (STA) scenarios often found with PowerShell GUIs.
    *   It's possible that some specific dialogs or UI elements might be Windows Forms-based, but the primary application structure and views appear to be WPF.

*   **No clear evidence of WinUI:**
    *   WinUI is a newer UI framework from Microsoft. While it also uses XAML, the namespaces and some controls differ from WPF. The namespaces seen (`http://schemas.microsoft.com/winfx/2006/xaml/presentation`) are classic WPF.
    *   There are no explicit references to WinUI-specific assemblies (like `Microsoft.UI.Xaml.dll`) or namespaces in the files we've examined.

**Conclusion:**

The front-end is **predominantly WPF**. It leverages Windows Forms for specific functionalities like `Application.DoEvents()` and potentially for some utility dialogs or interoperability needs, but the core UI framework and the way views are constructed and managed point strongly to WPF.

There is no direct evidence to suggest the use of WinUI in this project from the files reviewed.

Would you like me to examine `Core.psm1` to see if it reveals more about the UI interactions or confirms the use of any other UI elements? This is where `Start-CoreApp` is defined, which should load the `MainWindow.xaml`.

