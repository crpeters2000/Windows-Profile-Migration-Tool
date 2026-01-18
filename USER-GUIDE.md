# User Guide - Windows Profile Migration Tool

## Installation

The tool is a standalone portable application, but it now consists of **two required files**.

1. Download the latest release ZIP.
2. Extract **ALL** contents to a folder (e.g., `C:\Tools\ProfileMigration`).
3. Ensure both files are present:
   - `ProfileMigration.ps1`
   - `Functions.ps1`

### ⚠️ Critical Note
**You must keep `Functions.ps1` in the same folder as `ProfileMigration.ps1`.**
If you move the script to a USB drive or another computer, you must copy **BOTH** files.

## Running the Tool

1. Right-click `ProfileMigration.ps1`.
2. Select **Run with PowerShell**.
3. Accept any User Account Control (UAC) prompts (Administrator rights are required).

## Basic Usage

### Exporting a Profile (Old Computer)
1. **Launch the tool** as Administrator.
2. **Select the User** you want to migrate from the dropdown list.
   - The tool will calculate the profile size automatically.
3. Click the **Export** button.
4. **Choose a location** to save the migration file (e.g., USB drive or network share).
   - The file will be named `Export-username-Date.zip`.
5. Wait for the process to complete.
   - You can monitor progress in the status bar.
   - A detailed HTML report will open when finished.

### Importing a Profile (New Computer)
1. **Copy the migration file** (ZIP) to the new computer (or insert USB drive).
2. **Launch the tool** as Administrator.
3. Click the **Browse...** button (next to Import section).
4. Select your `.zip` migration file.
5. (Optional) **Modify the Target Username**:
   - By default, it imports to the same username.
   - To migrate to a different account, simply type the new username in the box.
   - Example: Exported from `jsmith`, Import to `john.smith`.
6. Click the **Import** button.
7. Wait for completion.
8. **Reboot** the computer when prompted.
9. Have the user log in!

---

## Advanced Features

### Domain Migration
To migrate a user to a domain account:
1. Ensure the computer is joined to the domain.
2. In the Import section, type the user as `DOMAIN\username`.
   - Example: `CONTOSO\jsmith`
3. Click Import.
4. You will be prompted for Domain Admin credentials to set up folder security.

### AzureAD / Entra ID Migration
To migrate to a cloud account:
1. Ensure the computer is AzureAD joined.
2. In the Import section, enter the email address (UPN).
   - Example: `jsmith@contoso.com`
3. Click Import.
4. The tool automatically handles the complex security settings for AzureAD.

### Profile Conversion (In-Place)
You can convert a profile on the *same* computer without exporting/importing (e.g., Joining a domain).
1. Click the purple **Convert** button in the header.
2. Select **Source Profile** (e.g., Local `jsmith`).
3. Select **Target Profile** (e.g., Domain `CONTOSO\jsmith`).
4. Click **Convert**.
5. The tool will rewrite the registry and permissions instanty.

### Profile Cleanup Wizard
When you start an export, the tool may show the Cleanup Wizard if it detects:
- Very large profile (>10 GB).
- Large "junk" folders (Downloads, Cache).
- Duplicate files.
You can use this wizard to select items to **exclude** from the export to reduce file size and time. This does **not** delete files from the source computer, it just skips them in the migration.

### Themes
Toggle between **Light Mode** and **Dark Mode** using the `L`/`D` button in the top right.

## Troubleshooting

- **Access Denied**: Ensure you are running as Administrator.
- **File In Use**: Ensure the user you are migrating is **Logged Off**. The tool will warn you if they are active.
- **Antivirus**: Some antivirus software may slow down the copy. You may need to temporarily pause protection if it blocks the script.
