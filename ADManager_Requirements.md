# AD Manager: Computer Mapper Requirements

This document tracks the technical requirements and feature roadmap for the Computer Mapper plugin, specifically the "Big Surgery" optimizations planned in April 2026.

## 1. Core Stability (Restored State)
- **Import**: Supports .csv and .xlsx (via ImportExcel module).
- **Matching**: Matches Excel TAG to AD Computer Name.
- **Identity**: Matches Excel 'Emp. ID' to AD User 'Description'.
- **UI**: 12-thread parallel update engine for AD modifications.

## 2. Feature Roadmap (To be implemented One-by-One)

### A. Advanced Identity Mapping [PENDING]
- **Priority 1**: Match 'Emp. ID' to AD 'Description'.
- **Priority 2**: If ID is missing, match 'Email' column to AD 'mail' property.
- **Priority 3**: If Email fails, strip the domain (@domain.com) and match prefix-to-SAMAccountName.
- **Back-population**: If AD data is found via Email, populate the 'Emp. ID' field in the export if it was blank.

### B. High-Performance Grid Engine [PENDING]
- **Bulk Swap**: Replace one-by-one `ObservableCollection.Add` with bulk attachment.
- **Grid Detachment**: Detach `ItemsSource` during heavy matching loops to prevent UI "Event Storms".
- **Async Updates**: Use `Dispatcher.BeginInvoke` carefully to prevent background thread deadlocks.

### C. Deep Discovery (Orphans) [PENDING]
- **Dedicated Button**: Add "🔍 Deep Scan" button separate from the main fetch.
- **Orphan Logic**: Identify Computers in AD but NOT in the file.
- **Stripping**: Automatically ignore Servers and Domain Controllers during discovery.
- **Performance**: Use set subtraction (HashSet) for instant orphan detection.

## 3. Technical Constraints
- **PowerShell 7**: Must maintain PS7 compatibility.
- **ADSI**: Use `[DirectoryServices.DirectoryEntry]` with `Secure, FastBind` for high-performance background threads.
- **Thread Safety**: All WPF property changes MUST happen on the UI thread via `Dispatcher`.
