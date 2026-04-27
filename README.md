# 🚀 iDM — Intelligent Directory Manager

iDM (Identity Data Manager) is a high-performance, modular PowerShell framework designed for enterprise-scale Active Directory operations. Built with a WPF modern interface and a robust plugin architecture, it allows administrators to perform bulk tasks with speed, safety, and precision.

---

## ✨ Key Features

-   **🛠️ Modular Plugin System**: Easily extend functionality by dropping plugin folders into the `Plugins/` directory.
-   **⚡ High-Performance Architecture**: 
    -   Multi-threaded background execution using PowerShell Runspaces.
    -   Optimized ADSI connectivity (`DirectoryEntry`) for rapid bulk updates.
    -   Parallelized update engine for mass record processing.
-   **🎨 Modern UI/UX**:
    -   WPF-based interactive dashboard.
    -   Dynamic tab-based navigation.
    -   Real-time status updates and progress tracking.
-   **📊 Enterprise Data Handling**:
    -   Native support for Excel (`.xlsx`) and CSV imports.
    -   Intelligent identity mapping (Matching TAGs to Computer Names, IDs to Descriptions).
-   **🛡️ Core Stability**:
    -   Comprehensive event-based connection management.
    -   Centralized logging system with rotation support.
    -   Sandboxed plugin loading to prevent core application crashes.

---

## 📂 Project Structure

```text
ADManager/
├── Launch.ps1              # Main Entry Point
├── Setup-ADManager.ps1     # Environment Configuration
├── Core/                   # Application Heart
│   ├── AppState.ps1        # Global State Management
│   ├── PluginLoader.ps1    # Dynamic Discovery Engine
│   ├── UIShell.ps1         # Main Interface Logic
│   └── Logger.ps1          # Diagnostic Logging
├── Plugins/                # Feature Modules
│   ├── UserManagement      # Account Lifecycle Management
│   ├── ComputerMapper      # Bulk Hardware/AD Synchronization
│   ├── OUMover             # Organizational Unit Relocation
│   └── SyncManager         # Identity Synchronization
└── UI/                     # Presentation Layer (XAML)
```

---

## 🧩 Featured Plugins

### 💻 Computer Mapper
Automates the synchronization between inventory reports and Active Directory. Supports deep discovery of orphaned objects and high-speed parallel updates.

### 👤 User Management
A comprehensive tool for user account maintenance, password resets, and account status auditing.

### 📁 OU Mover
Safely transition Active Directory objects across OU hierarchies with validation and logging.

---

## 🚀 Getting Started

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/miirfan84/ADManager.git
   ```
2. **Setup Dependencies**:
   Run the setup script to ensure all required modules (like `ImportExcel`) are available:
   ```powershell
   .\Setup-ADManager.ps1
   ```
3. **Launch the App**:
   ```powershell
   .\Launch.ps1
   ```

---

## 🛠️ Developer Guide (Plugin Authoring)

iDM is built to be extended. To create a new plugin:
1. Create a folder in `Plugins/`.
2. Define your metadata in `plugin.json`.
3. Design your UI in `Tab.xaml`.
4. Wire your events in `Handlers.ps1`.

Refer to the internal [Plugin Documentation](PLUGINS_README.md) for full API details.

---

## 📜 Requirements

-   **OS**: Windows 10/11 or Windows Server.
-   **PowerShell**: Version 7.2+ recommended (v5.1 compatible).
-   **Permissions**: Active Directory RSAT tools and appropriate domain permissions.

---

*Built with ❤️ for Active Directory Administrators.*
