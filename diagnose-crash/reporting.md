# Reporting a Crash Upstream

Read this after diagnosing a crash and concluding it is an upstream application or library bug.

---

## 1. Verify Scope

Before filing an issue with an open-source project or software vendor:
- Is this a bug in the application itself, or a conflict caused by a third-party antivirus, shell extension, or out-of-date GPU driver?
- Check if running the application in a clean state (e.g. without plugins or in Safe Mode) reproduces the issue.

---

## 2. Gather the Crash Packet

A high-quality bug report should include:
1. **Windows Version & Build**:
   ```powershell
   [System.Environment]::OSVersion.VersionString
   (Get-CimInstance Win32_OperatingSystem).Caption + " Build " + (Get-CimInstance Win32_OperatingSystem).BuildNumber
   ```
2. **Exact Application Version & Architecture**: x64, ARM64, or x86.
3. **Event Viewer Entry**:
   Copy the text from Event ID 1000:
   - Faulting application name, version, timestamp
   - Faulting module name, version, timestamp
   - Exception code (e.g. `0xc0000005`)
   - Fault offset
4. **Steps to Reproduce**: Minimal, deterministic steps that trigger the crash.
5. **Sanitizing Crash Dumps**:
   - Never post full memory dumps (`.dmp`) publicly on GitHub or public forums without sanitization: memory dumps can contain passwords, tokens, API keys, and personal data.
   - If a minidump is requested by maintainers, confirm it does not contain sensitive data or provide stack backtrace text instead.

---

## 3. Search Before Filing

Avoid duplicate issue submissions:
- Search the vendor's GitHub repository or issue tracker using the **faulting module**, **exception code**, and **distinctive function names** from the stack.
- Check both open and closed issues. A matching closed issue that reproduces on the current build is a regression report.
