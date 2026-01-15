# 📄 CCWIS File Preview PCF Control

A high-fidelity **Power Apps Component Framework (PCF)** control that allows users to **preview attachments directly inside Dynamics 365 / Power Apps** without downloading them.

This control solves common browser limitations by using **Magic Byte Detection** to correctly identify and render files, even if the server returns generic MIME types (like `application/octet-stream`) or missing file extensions.

## ✨ Key Features

* **True Fullscreen Popup:** The preview modal attaches to the main window body, ensuring a full-screen experience that isn't confined by the form field's size.
* **High-Fidelity Rendering:**
  * **Word (.docx):** Renders actual pages with correct fonts, layout, and pagination (via `docx-preview`).
  * **Excel (.xlsx, .csv):** Renders a scrollable grid with proper headers and gridlines.
  * **Text (.txt):** Automatically detects text files even without extensions using binary scanning.
  * **PDF & Images:** Native browser rendering.
* **Smart Detection:** Uses **Magic Byte (File Signature) analysis** to ignore incorrect server labels and render files based on their actual binary content.
* **User Friendly Error States:**
  * **Empty State:** Displays "Please attach a file" if the field is empty.
  * **Unsupported:** Displays a clear "Format not supported" message without forcing a download.

## 🚀 Supported Formats

| Format | Tech Used | Behavior |
| :--- | :--- | :--- |
| **PDF** | Native | Opens natively in iframe. |
| **Word** (`.docx`) | `docx-preview` | Renders as a document with page breaks & styles. |
| **Excel** (`.xlsx`, `.csv`) | `SheetJS` (`xlsx`) | Renders as a styled HTML table with scrolling. |
| **Images** (`png`, `jpg`, `gif`) | Native | Renders scaled to fit the modal. |
| **Text** (`.txt`) | TextDecoder | Renders inside a scrollable code block. |

## 🛠️ Installation & Configuration

### **1. Download & Import**
Download the **Unmanaged Solution** (`.zip`) file provided in this repository (or Release section).

1. Go to **make.powerapps.com** -> **Solutions**.
2. Click **Import Solution**.
3. Select the `FilePreviewPopUp.zip` file you downloaded.
4. Proceed with the import (Publish all customizations after import).

### **2. Add to Form**
1. Open the **Form Editor** (Classic or Modern).
2. Add a **Text Field** to your form (this acts as a placeholder for the button).
3. Go to **Components** -> **Add Component** -> Select **CCWIS Gold Preview**.
4. Configure the properties:

| Property | Description |
| :--- | :--- |
| **Dummy Field** | Bind this to the text field you added (required by PCF). |
| **Target File Col Name** | **CRITICAL:** Enter the **Logical Name** of the file column you want to preview (e.g., `cw_attach` or `cr56_fileupload`). Do not use the Display Name. |

5. **Save & Publish**.

## 💻 Build from Source

If you want to modify the code, follow these steps:

**Prerequisites:**
* Node.js & npm
* Microsoft Power Platform CLI (`pac`)
* .NET SDK

**Build Commands:**
```bash
# 1. Install Dependencies
npm install

# 2. Build the Control
npm run build

# 3. Package the Solution
cd FilePreviewPopUp
dotnet build --configuration Release /p:SolutionPackageType=Unmanaged
