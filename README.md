# LinkForge

**LinkForge** is a Windows-based desktop application built with **Tkinter** that allows users to create junction folders using the `mklink` command. It provides an easy-to-use interface for creating and managing symbolic links between folders. Additionally, the app keeps track of all the junctions created, displaying them in a **History Panel** for easy reference.

## Features
- **Create Junction Folders:** Easily create junction folders by specifying a source location, target location, and a junction folder name.
- **History Panel:** Tracks all junctions created by the application, displaying them in a user-friendly interface.
- **Windows Integration:** Uses the `mklink` command to create junctions, a feature native to Windows.
- **Intuitive Tkinter Interface:** A simple and clean interface built with Tkinter for ease of use.

## Installation

### 1. Using the Executable (.exe)
   You can download the latest **LinkForge** `.exe` release from the [releases section](#). Once downloaded:
   - Simply double-click the `.exe` file to launch the application.
   
### 2. From Source (Python)
   If you prefer to run the source code directly, follow these steps:

   1. Clone the repository:
      ```bash
      git clone https://github.com/Mathew005/LinkForge.git
      ```
   2. Navigate to the project directory:
      ```bash
      cd LinkForge
      ```
   3. Run the application:
      ```bash
      python main.py
      ```

## How to Use

1. **Select Source Folder:** Choose the folder you want to link as the source.
2. **Choose Target Folder:** Pick the target location where the junction folder will be created.
3. **Enter Junction Name:** Provide a name for the junction folder.
4. **Create Junction:** Click the button to create the junction using the `mklink` command.
5. **View History:** All created junctions will be listed in the **History Panel**. You can track all your previous junctions here.

## Example
Here is an example of how the **LinkForge** app will look:
1. **Source Folder:** `C:\Users\User\Documents\ImportantFolder`
2. **Target Folder:** `D:\Backup\ImportantLink`
3. **Junction Name:** `ImportantFolderLink`
   - A junction will be created at `D:\Backup\ImportantLink` pointing to `C:\Users\User\Documents\ImportantFolder`.

## History Panel
The **History Panel** keeps track of all junctions created by **LinkForge**. Each entry includes:
- **Junction Name**
- **Source Location**
- **Target Location**
- **Date Created**

You can easily review your previous junctions here.

## Compiling the Application into an .exe
If you'd like to compile the Python source into an executable yourself:
1. Install **PyInstaller**:
   ```bash
   pip install pyinstaller
