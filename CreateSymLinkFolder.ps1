$source = @'
using System;
using System.Reflection;
using System.Windows.Forms;
public class MyFolderSelector
{
    public string DefaultPath, Title, Message;
    public MyFolderSelector(string defaultPath = "MyComputer", string title = "Select a folder", string message = "")
    {
        DefaultPath = defaultPath;
        Title = title;
        Message = message;
    }
    public string GetPath()
    {
        // filter for which methods we want to get
        BindingFlags c_flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        // get the assembly of the file dialog so we can access its internal types
        var fileDialogAssembly = typeof(FileDialog).Assembly;
        // get the iFile so we can use its functions to put our settings and show our dialog
        var iFileDialog = fileDialogAssembly.GetType("System.Windows.Forms.FileDialogNative+IFileDialog");
        var openFileDialog = new OpenFileDialog
        { // our dialog options
            AddExtension = false,
            CheckFileExists = false,
            DereferenceLinks = true,
            Filter = "Folders|\n",
            InitialDirectory = DefaultPath,
            Multiselect = false,
            Title = Title
        };
        // create vista dialog 
        var ourDialog = (typeof(OpenFileDialog).GetMethod("CreateVistaDialog", c_flags)).Invoke(openFileDialog, new object[] { });
        // attatch our options
        typeof(OpenFileDialog).GetMethod("OnBeforeVistaDialog", c_flags).Invoke(openFileDialog, new[] { ourDialog });
        iFileDialog.GetMethod("SetOptions", c_flags).Invoke(ourDialog, new object[] { (uint)(0x20000840 | 0x00000020 /*defaultOptions | folderPickerOption*/) });
        // generate update event constructor so the file name updates after selection
        var vistaDialogEventsConstructorInfo = fileDialogAssembly
        .GetType("System.Windows.Forms.FileDialog+VistaDialogEvents")
        .GetConstructor(c_flags, null, new[] { typeof(FileDialog) }, null);
        var adviseParametersWithOutputConnectionToken = new[] { vistaDialogEventsConstructorInfo.Invoke(new object[] { openFileDialog }), 0U };
        // add the event to the dialog
        iFileDialog.GetMethod("Advise").Invoke(ourDialog, adviseParametersWithOutputConnectionToken);
        // show the dialog to user and get response
        int retVal = (int)iFileDialog.GetMethod("Show").Invoke(ourDialog, new object[] { IntPtr.Zero });
        // returns the file name if the user made a selection
        return retVal == 0 ? openFileDialog.FileName : "";
    }
}
'@
# Function to check for elevated privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Confirm-Overwrite {
    param ($linkPath)
    Add-Type -AssemblyName PresentationFramework

    # Presume that it's not a folder
    $isDirectory = $false

    if (-not [string]::IsNullOrWhiteSpace($linkPath)) {
        # If the input box isn't empty, check if it's a valid path
        if (Test-Path $linkPath -PathType Leaf) {
            # if the path points to a file, it is not a folder
            $isDirectory = $false
        } elseif (Test-Path $linkPath -PathType Container) {
            # if the path points to a folder, it is a folder
            $isDirectory = $true
        }
    }

    $message = "A $(if ($isDirectory) { 'folder' } else { 'file' }) with the name '$(Split-Path $linkPath -Leaf)' already exists. Overwrite?"
    $result = [System.Windows.MessageBox]::Show(
        $message,
        "Confirm Overwrite",
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Warning
    )
    return $result
}

function Escape-WildcardCharacters {
    param ([string]$path)
    $escapedPath = $path -replace '([*?\[\]])', '``$1'
    return $escapedPath
}

# Save the initial working directory
$initialWorkingDirectory = [System.IO.Directory]::GetCurrentDirectory()
$escapedInitialWorkingDirectory = Escape-WildcardCharacters -path $initialWorkingDirectory

if (-not (Test-Admin)) {
    Write-Output "Please run this script as an administrator. Let's try rerunning automatically..."
    
    # Re-run the script with elevated privileges, passing the initial working directory as an argument
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$escapedInitialWorkingDirectory`"" -Verb RunAs
    Exit  # Exit the non-elevated instance
}

# Retrieve the working directory from arguments
$originalWorkingDirectory = $args[0]
Set-Location -Path $originalWorkingDirectory

# Function to show folder picker dialog
function Show-FolderPickerDialog {

    Add-Type -Language CSharp -TypeDefinition $source -ReferencedAssemblies ("System.Windows.Forms")

    $path = $originalWorkingDirectory
    $title = "Select Folder"
    $description = "Select a folder to create a symbolic link"

    $out = [MyFolderSelector]::new($path, $title, $description).GetPath()
    
    if ($out -eq "") {
        return $null
    } else {
        return $out
    }
}

# Main script
$selectedPath = Show-FolderPickerDialog

if ($selectedPath) {
    $currentDir = Get-Location -PSProvider FileSystem
    $linkName = Split-Path -Leaf $selectedPath
    $linkPath = Join-Path -Path $currentDir -ChildPath $linkName

    # Escape wildcard characters in the paths
    $escapedLinkPath = Escape-WildcardCharacters -path $linkPath
    $escapedSelectedPath = Escape-WildcardCharacters -path $selectedPath

    Write-Output "Original Working Directory: $originalWorkingDirectory"
    Write-Output "Current Directory: $currentDir"
    Write-Output "Selected Path: $selectedPath"
    Write-Output "Escaped Selected Path: $escapedSelectedPath"
    Write-Output "Link Path: $linkPath"
    Write-Output "Escaped Link Path: $escapedLinkPath"

    if (Test-Path -Path $escapedSelectedPath -PathType Container) {
        if (Test-Path -Path $escapedLinkPath) {
            $choice = Confirm-Overwrite -linkPath $escapedLinkPath
            if ($choice -eq 'Yes') {
                # Remove the existing symlink
                Remove-Item -Path $escapedLinkPath -Force
                # Create a new directory symbolic link
                New-Item -ItemType SymbolicLink -Path $linkPath -Target $escapedSelectedPath -ErrorAction Stop
                Write-Output "Overwritten directory symbolic link: $escapedLinkPath -> $escapedSelectedPath"
            } elseif ($choice -eq 'No') {
                Write-Output "Symbolic link (or file/folder with that name) already existed! Operation canceled by user by not overwriting."
            } else {
                Write-Output "Operation canceled by user."
            }
        } else {
            # Create a new directory symbolic link
            New-Item -ItemType SymbolicLink -Path $linkPath -Target $escapedSelectedPath -ErrorAction Stop
            Write-Output "Created directory symbolic link: $escapedLinkPath -> $escapedSelectedPath"
        }
    } else {
        Write-Output "The selected path is not a folder."
    }
} else {
    Write-Output "No folder was selected."
}

# Keep the window open
#Read-Host -Prompt "Press Enter to exit"
Write-Output "`n`nThis window will autoclose soon"
Start-Sleep -Seconds 1