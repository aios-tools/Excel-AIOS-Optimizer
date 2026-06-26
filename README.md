# AIOS-Optimizer Installation & Setup Guide

This guide explains how to install and enable the **AIOS-Optimizer** Excel Add-in.

---

## Installation Steps

### Step 1: Download the Release Archive
1. Go to the **Releases** section of the repository on GitHub.
2. Download the latest release ZIP file (e.g., `AiosOptimizer.zip`).

### Step 2: Copy Files to Excel Add-ins Directory
1. Extract the contents of the downloaded ZIP file.
2. Open Windows File Explorer.
3. Paste the following path into the File Explorer address bar and press **Enter** to open the Excel Add-ins folder:
   ```cmd
   %appdata%\Microsoft\AddIns
   ```
4. Copy the extracted files and paste them directly into this folder.

### Step 3: Enable the Add-in in Microsoft Excel
1. Open **Microsoft Excel**.
2. Navigate to **File** > **Options** (located at the bottom of the left sidebar).
3. In the Excel Options window, select **Add-ins** from the left-hand menu.
4. At the bottom of the window, locate the **Manage** drop-down list. Ensure **Excel Add-ins** is selected and click the **Go...** button.
5. In the Add-ins dialog:
   - Check the box next to the **AIOS-Optimizer** add-in macro.
   - *Note: If it does not appear in the list, click **Browse...**, navigate to your `%appdata%\Microsoft\AddIns` folder, select the add-in file, and click **OK**.*
6. Click **OK** to apply the changes.

Once completed, the **AIOS Optimizer** tab will appear on your Excel Ribbon, confirming that the add-in is enabled and ready to use.

---

> [!NOTE]
> For more details on using the optimizer, please refer to the [AIOS User Manual](AIOS_Manual_Usuario.md) in the same directory.
