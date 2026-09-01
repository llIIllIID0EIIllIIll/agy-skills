---
name: diagnose-crash
description: >
  Diagnose why an application or process crashed, threw an unhandled exception, froze,
  or caused a BSOD on Windows. Use when an application terminates unexpectedly, crashes,
  when asked "why did X crash", "X keeps crashing", or when investigating Windows Event
  Viewer error logs, Windows Error Reporting (WER), or crash dumps (.dmp).
  Triggers: crash, exception, 0xC0000005, access violation, Event ID 1000, WER,
  Windows Error Reporting, minidump, .dmp, procdump, cdb, windbg, BSOD, BugCheck.
  Covers crash root cause analysis, stack backtraces, and reporting.
---

# Diagnosing a Crash on Windows

Work from evidence. The goal is an honest, verifiable account of what happened, not a plausible-sounding story.

---

## 1. Establish the Facts (Windows Event Log)

The Windows Event Log is the first and fastest source of crash evidence.

### Application Crashes (Event ID 1000 & 1001)

Run this in PowerShell to inspect recent application crashes:

```powershell
# Get recent application crashes (Event ID 1000)
Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1000} -MaxEvents 5 -ErrorAction SilentlyContinue |
    Format-List TimeCreated, Id, Message

# Get Windows Error Reporting (WER) events (Event ID 1001)
Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1001} -MaxEvents 5 -ErrorAction SilentlyContinue |
    Format-List TimeCreated, Id, Message
```

In the **Event ID 1000** payload, extract these critical fields:
- **Faulting application name & version**: The exact binary that crashed.
- **Faulting module name & version**: The DLL or library where the failure occurred (e.g., `ntdll.dll`, `d3d11.dll`, `ucrtbase.dll`).
- **Exception code**: The NTSTATUS hex code (e.g., `0xc0000005`). See the Exception Code Reference below.
- **Fault offset**: The instruction offset within the faulting module (e.g., `0x000000000001a4e2`).
- **Process ID & Application start time**: To correlate with lifetime and resources.

### Application Hangs / Freezes (Event ID 1002)

```powershell
Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1002} -MaxEvents 5 -ErrorAction SilentlyContinue |
    Format-List TimeCreated, Message
```

### System Crashes / BSOD (System Event ID 1001 / BugCheck)

```powershell
Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-WER-SystemErrorReporting'} -MaxEvents 3 -ErrorAction SilentlyContinue |
    Format-List TimeCreated, Message
```

---

## 2. Rule Out the Boring Causes First

Before blaming application logic or compiler bugs, rule out system resource starvation:

```powershell
# Check available physical and virtual memory
Get-CimInstance Win32_OperatingSystem | Select-Object @{N='FreeRAM_GB';E={[math]::Round($_.FreePhysicalMemory/1MB,2)}},
    @{N='TotalRAM_GB';E={[math]::Round($_.TotalVisibleMemorySize/1MB,2)}},
    @{N='FreeVirtual_GB';E={[math]::Round($_.FreeVirtualMemory/1MB,2)}},
    @{N='TotalVirtual_GB';E={[math]::Round($_.TotalVirtualMemorySize/1MB,2)}}

# Check disk space on system drive
Get-PSDrive -PSProvider FileSystem | Select-Object Name, @{N='Free_GB';E={[math]::Round($_.Free/1GB,2)}}, @{N='Total_GB';E={[math]::Round($_.Used/1GB + $_.Free/1GB,2)}}

# Check top memory-consuming processes
Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 ProcessName, Id, @{N='WS_MB';E={[math]::Round($_.WorkingSet64/1MB,2)}}
```

A process killed due to system pagefile exhaustion (Commit Limit) is not an application logic bug.

---

## 3. Correlate Against the Timeline

The crash timestamp is vital. Cross-reference it with:

1. **Reliability Monitor Records**:
   ```powershell
   Get-CimInstance -ClassName Win32_ReliabilityRecords -Filter "TimeGenerated >= '$((Get-Date).AddDays(-2).ToString('yyyy-MM-dd'))'" |
       Select-Object TimeGenerated, EventIdentifier, SourceName, Message | Format-Table -Wrap
   ```
2. **Recent Windows Updates and Drivers**:
   ```powershell
   Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 5
   ```
3. **Filesystem Timestamps**: Look at the directory the program operates on. Files modified on the exact second of the crash often reveal the input file that triggered the crash.

---

## 4. Analyze Crash Dumps (.dmp)

Windows can capture full or mini crash dumps automatically.

### Locating Crash Dumps

Check these standard locations:
- **User-mode WER Crash Dumps**: `%LOCALAPPDATA%\CrashDumps\` (e.g. `C:\Users\<user>\AppData\Local\CrashDumps`)
- **WER Report Archives**: `%PROGRAMDATA%\Microsoft\Windows\WER\ReportArchive\`
- **Kernel Minidumps (BSOD)**: `C:\Windows\Minidump\`
- **Full Memory Dump**: `C:\Windows\MEMORY.DMP`

```powershell
# List user-mode crash dumps
Get-ChildItem -Path "$env:LOCALAPPDATA\CrashDumps" -Filter "*.dmp" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object Name, Length, LastWriteTime
```

### Enabling Local Dumps (If none were captured)

If Windows is not saving user-mode crash dumps for the crashing app:

```powershell
# Enable LocalDumps via Registry (Run in elevated PowerShell)
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps"
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "DumpType" -Value 2 -Type DWord           # 1 = Mini, 2 = Full
Set-ItemProperty -Path $regPath -Name "DumpCount" -Value 5 -Type DWord
Set-ItemProperty -Path $regPath -Name "DumpFolder" -Value "%LOCALAPPDATA%\CrashDumps" -Type ExpandString
```

### Symbolizing and Inspecting Dumps with WinDbg / CDB

If Windows SDK debugging tools (`cdb.exe` or `windbg.exe`) are available:

```cmd
set _NT_SYMBOL_PATH=srv*https://msdl.microsoft.com/download/symbols
cdb -z "C:\path\to\crash.dmp" -c "!analyze -v; q"
```

In the output, note:
- `FAULTING_IP`: Instruction pointer at crash.
- `EXCEPTION_RECORD`: Code and memory target address.
- `STACK_TEXT`: The thread backtrace.
- Third-party DLLs in the loaded module list.

---

## 5. Common Windows Exception Codes Reference

| Exception Code | Constant Name | Common Root Cause |
| :--- | :--- | :--- |
| **`0xC0000005`** | `STATUS_ACCESS_VIOLATION` | Null pointer dereference, invalid memory pointer, buffer overflow, read/write to unmapped address. |
| **`0xC00000FD`** | `STATUS_STACK_OVERFLOW` | Infinite recursion or huge local array allocated on stack. |
| **`0xC0000409`** | `STATUS_STACK_BUFFER_OVERRUN` | Security buffer overrun (/GS fail-fast triggered), aborting before malicious execution. |
| **`0x80000003`** | `STATUS_BREAKPOINT` | Hardcoded `__debugbreak()` / `int 3`, or an assert() condition failed. |
| **`0xC0000022`** | `STATUS_ACCESS_DENIED` | Permission or security token failure (file, service, or registry access blocked). |
| **`0xE06D7363`** | Microsoft C++ Exception | Unhandled C++ `throw` statement. Often has parameter pointing to the thrown C++ object type. |
| **`0xC0000374`** | `STATUS_HEAP_CORRUPTION` | Double free, heap buffer overflow, corrupted malloc/heap chunk. |

---

## 6. Diagnosis Report Format

When delivering diagnosis results to the user:

1. **What crashed**: Binary name, PID, version, and what it was doing at the time.
2. **Mechanism**:
   - Faulting module and offset.
   - Exception code and human explanation.
   - Separate what the evidence **proves** from what you are **inferring**.
3. **Data Loss**: Did the crash corrupt files? Check `%TEMP%`, autosave locations, or Recycle Bin before concluding data is gone.
4. **Recurrence & Mitigation**: Is this reproducible? How to prevent it (patch, configuration change, disabling faulty plugin/extension).
5. **Clean Up**: Delete any temporary dump copies you extracted to avoid filling disk space or leaking sensitive process memory.
