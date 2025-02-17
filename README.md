# 📂 Automate Folder & File Upload to SharePoint Online and Extracts page count for DOCX and PPTX files and updates metadata

## 🚀 Overview
This PowerShell script automates the **creation of hierarchical folder structures** and **uploads PDF, Word (DOCX), and PowerPoint (PPTX) files** to a SharePoint Online **document library**. It also updates metadata (Page Count) for DOCX and PPTX files. If the document library does not exist, it will be created automatically.

## 🎯 Features
✅ Connects to **SharePoint Online** using **PnP PowerShell**  
✅ **Automatically creates** document libraries (if they don’t exist)  
✅ **Preserves folder hierarchy** during upload  
✅ **Uploads files** (PDF, DOCX, PPTX) to the correct SharePoint folders  
✅ **Prevents duplicate folders** by checking existence before creation  
✅ **Extracts page count** for DOCX and PPTX files and updates metadata  
✅ **Provides console logs** for real-time tracking  

## 📋 Prerequisites
Before running this script, ensure you have:

1️⃣ **PnP PowerShell module installed:**  
   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser
   ```
2️⃣ **Valid SharePoint Online credentials** with appropriate permissions  
3️⃣ **PowerShell execution policy set to allow script execution:**  
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope Process
   ```
4️⃣ **Enable COM objects** for extracting page counts from Word & PowerPoint files (Optional)  

## 🛠️ How to Use

1️⃣ **Update the script** with your SharePoint details:
   ```powershell
   $siteUrl = "https://yoursharepointdomain/sites/YourSiteName"
   $libraryName = "YourLibraryName"
   $libraryDisplayName = "Your Library Display Name"
   $localPath = "C:\Path\To\Your\Local\Folder"
   ```

2️⃣ **Run the script** in PowerShell:
   ```powershell
   .\YourScriptName.ps1
   ```

3️⃣ **Monitor the logs** for progress updates.

## 📂 Example Folder Structure
### 🎯 Local Folder Structure:
```
📁 LocalFolder
 ├── 📁 A1
 │   ├── 📄 file1.pdf
 │   ├── 📄 file2.docx
 │   ├── 📁 B1
 │       ├── 📄 file3.pptx
 │       ├── 📁 C1
 │           ├── 📄 file4.pdf
```
### 🔄 SharePoint Structure After Upload:
```
📂 SharePoint Library
 ├── 📁 A1
 │   ├── 📄 file1.pdf
 │   ├── 📄 file2.docx
 │   ├── 📁 B1
 │       ├── 📄 file3.pptx
 │       ├── 📁 C1
 │           ├── 📄 file4.pdf
```

## ⚠️ Important Notes
🔹 This script **does not overwrite** existing files.  
🔹 If metadata updates (like Page Count for DOCX/PPTX) are needed, ensure COM objects are enabled.  
🔹 **Execution policy** must allow script execution (`Set-ExecutionPolicy RemoteSigned`).  
🔹 **Admin access** may be required for running PnP PowerShell commands.  

## 🏆 Benefits
🚀 Saves time by **automating folder creation and file uploads**  
📂 Maintains a **structured and organized SharePoint library**  
⚡ Works efficiently with **large datasets and complex hierarchies**  
📊 Enhances document metadata with **Page Count updates**  

## 📞 Need Help?
If you encounter any issues, feel free to reach out! 💡 Happy Automating! 🚀

