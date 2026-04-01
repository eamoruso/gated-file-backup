<<<<<<< HEAD
﻿// ----------------------------------------------------------
// 
// Developed by: Dr. Edward Amoruso
// University of Central Florida
// 
// ----------------------------------------------------------

#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#define SOFTWARE_VERSION    7.9
=======
﻿// FileMonitorBackup6.cpp
//
// Patched: Asynchronous monitoring, structure checks, white-list,
// entropy probes, hardened/hidden backup repo with GUID names + 
// manifest, JSON logging.
//
// Build: /std:c++17 ; Link: Ole32.lib; Advapi32.lib; Aclapi.lib;
//       Shlwapi.lib; Shell32.lib
>>>>>>> b8978569a7f1ec61f96c44f07beb82d600692e5c

#include <windows.h>
#include <aclapi.h>
#include <shlobj.h>
#include <sddl.h>
#include <shlwapi.h>
#include <tlhelp32.h>   // Added for process checking
#include <psapi.h>      // For GetMappedFileName
#include <winternl.h>   // For Process Identification

#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <array>
#include <algorithm>
#include <random>
#include <cmath>
#include <thread>
#include <chrono>

// Entropy constants for different file types
#define GENERAL_ENTROPY_VALUE   7.7     // Default entropy value for most files
#define DOCX_ENTROPY_VALUE      7.6     // Entropy for DOCX due to compression
#define XLSX_ENTROPY_VALUE      7.6     // Entropy for XLSX due to compression
#define PDF_ENTROPY_VALUE       7.8     // Entropy for PDFs mixed binary/text content
#define JPG_ENTROPY_VALUE       7.9     // Entropy for JPEG images relatively high

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Psapi.lib")

static const int   RETRY_MAX_ATTEMPTS = 5;      // R/W how many times we will retry
static const DWORD RETRY_INITIAL_DELAY = 50;    // R/W ms – first sleep
static const DWORD RETRY_MAX_DELAY = 800;       // R/W ms – upper bound for back‑off

using namespace std;

//  Global shutdown
static HANDLE g_shutdownEvent = nullptr;
BOOL WINAPI ConsoleCtrlHandler(DWORD type)
{
    if (type == CTRL_C_EVENT ||
        type == CTRL_BREAK_EVENT ||
        type == CTRL_CLOSE_EVENT ||
        type == CTRL_SHUTDOWN_EVENT)
    {
        if (g_shutdownEvent)
            SetEvent(g_shutdownEvent);
        return TRUE;
    }
    return FALSE;
}

// UTF helpers
string ToUtf8(const wstring& w)
{
    if (w.empty()) return {};
    int sz = WideCharToMultiByte(CP_UTF8, 0,
        w.data(), (int)w.size(),
        nullptr, 0, nullptr, nullptr);
    string out(sz, '\0');
    WideCharToMultiByte(CP_UTF8, 0,
        w.data(), (int)w.size(),
        &out[0], sz, nullptr, nullptr);
    return out;
}

// Backup helpers - function for process integrity checking
bool IsHighIntegrityProcess()
{
    HANDLE hToken = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken))
        return false;

    TOKEN_ELEVATION elevation;
    DWORD dwSize = sizeof(elevation);
    BOOL bResult = GetTokenInformation(hToken, TokenElevation, &elevation, sizeof(elevation), &dwSize);

    CloseHandle(hToken);
    return bResult && elevation.TokenIsElevated;
}

bool ValidateAccessingProcess(DWORD processId)
{
    if (!processId) return false;

    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
    if (!hProcess) return false;

    // Check if it's our own process (legitimate backup)
    if (processId == GetCurrentProcessId())
    {
        CloseHandle(hProcess);
        return true;
    }
    CloseHandle(hProcess);
    return false;
}

wstring ToWide(const string& s)
{
    if (s.empty()) return {};
    int sz = MultiByteToWideChar(CP_UTF8, 0,
        s.data(), (int)s.size(),
        nullptr, 0);
    wstring out(sz, L'\0');
    MultiByteToWideChar(CP_UTF8, 0,
        s.data(), (int)s.size(),
        &out[0], sz);
    return out;
}

//   GUID Function
wstring GuidString()
{
    GUID g{};
    if (FAILED(CoCreateGuid(&g))) terminate();
    wchar_t buf[64];
    StringFromGUID2(g, buf, ARRAYSIZE(buf));
    return wstring(buf + 1, wcslen(buf) - 2);
}

// Logging Function
void AppendJson(const wstring& path, const string& line)
{
    ofstream f(ToUtf8(path), ios::app | ios::binary);
    if (f) f << line << "\n";
}

// Get Document Paths
wstring GetDocumentsPath()
{
    PWSTR p = nullptr;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, 0, nullptr, &p)))
    {
        wstring out(p);
        CoTaskMemFree(p);
        return out;
    }
    return L"C:\\Users\\Public\\Documents";
}

// Enhanced ACL hardening
bool HardenFolderDACL(const wstring& dir)
{
    // Ensure directory exists
    DWORD dwAttrib = GetFileAttributesW(dir.c_str());
    if (dwAttrib == INVALID_FILE_ATTRIBUTES)
    {
        if (!CreateDirectoryW(dir.c_str(), nullptr))
        {
            wcout << L"Failed to create directory: " << dir << L"\n";
            return false;
        }
    }

    // Simple SDDL string that works on Windows 10
    LPCWSTR sddl = L"D:P(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200A9;;;OW)";

    PSECURITY_DESCRIPTOR pSD = nullptr;
    ULONG ulSDSize = 0;

    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
        sddl, SDDL_REVISION_1, &pSD, &ulSDSize))
    {
        DWORD error = GetLastError();
        wcout << L"ConvertStringSecurityDescriptor failed: " << error << L"\n";
        return false;
    }

    PACL pAcl = nullptr;
    BOOL bDaclPresent = FALSE;
    BOOL bDaclDefaulted = FALSE;

    if (!GetSecurityDescriptorDacl(pSD, &bDaclPresent, &pAcl, &bDaclDefaulted))
    {
        DWORD error = GetLastError();
        wcout << L"GetSecurityDescriptorDacl failed: " << error << L"\n";
        LocalFree(pSD);
        return false;
    }

    if (!bDaclPresent || pAcl == nullptr)
    {
        wcout << L"No DACL present in security descriptor\n";
        LocalFree(pSD);
        return false;
    }

    DWORD dwRes = SetNamedSecurityInfoW(
        const_cast<LPWSTR>(dir.c_str()),
        SE_FILE_OBJECT,
        DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION,
        nullptr, nullptr, pAcl, nullptr);

    LocalFree(pSD);

    if (dwRes != ERROR_SUCCESS)
    {
        wcout << L"SetNamedSecurityInfo failed: " << dwRes << L"\n";
        return false;
    }

    return true;
}

static bool IsWordTempFile(const wstring& n)
{
    wstring l = n;
    transform(l.begin(), l.end(), l.begin(), towlower);
    return l.starts_with(L"~$") ||
        l.ends_with(L".tmp") ||
        l.ends_with(L".asd") ||
        l.ends_with(L".wbk") ||
        l.ends_with(L".~tmp");
}

bool IsSupportedFileType(const wstring& name)
{
    wstring l = name;
    transform(l.begin(), l.end(), l.begin(), towlower);

    // Only allow these specific file types
    return l.ends_with(L".doc") ||
        l.ends_with(L".docx") ||
        l.ends_with(L".docm") ||
        l.ends_with(L".dot") ||
        l.ends_with(L".dotx") ||
        l.ends_with(L".dotm") ||
        l.ends_with(L".xls") ||
        l.ends_with(L".xlsx") ||
        l.ends_with(L".xlsm") ||
        l.ends_with(L".xlsb") ||
        l.ends_with(L".xlt") ||
        l.ends_with(L".xltx") ||
        l.ends_with(L".xltm") ||
        l.ends_with(L".pdf") ||
        l.ends_with(L".jpg") ||
        l.ends_with(L".jpeg");
}

bool IsOfficeFile(const wstring& name)
{
    wstring l = name;
    transform(l.begin(), l.end(), l.begin(), towlower);
    return l.ends_with(L".doc") ||
        l.ends_with(L".docx") ||
        l.ends_with(L".docm") ||
        l.ends_with(L".dot") ||
        l.ends_with(L".dotx") ||
        l.ends_with(L".dotm") ||
        l.ends_with(L".xls") ||
        l.ends_with(L".xlsx") ||
        l.ends_with(L".xlsm") ||
        l.ends_with(L".xlsb") ||
        l.ends_with(L".xlt") ||
        l.ends_with(L".xltx") ||
        l.ends_with(L".xltm");
}

bool IsWordFile(const wstring& name)
{
    wstring l = name;
    transform(l.begin(), l.end(), l.begin(), towlower);
    return l.ends_with(L".doc") ||
        l.ends_with(L".docx") ||
        l.ends_with(L".docm") ||
        l.ends_with(L".dot") ||
        l.ends_with(L".dotx") ||
        l.ends_with(L".dotm");
}

bool IsExcelFile(const wstring& name)
{
    wstring l = name;
    transform(l.begin(), l.end(), l.begin(), towlower);
    return l.ends_with(L".xls") ||
        l.ends_with(L".xlsx") ||
        l.ends_with(L".xlsm") ||
        l.ends_with(L".xlsb") ||
        l.ends_with(L".xlt") ||
        l.ends_with(L".xltx") ||
        l.ends_with(L".xltm");
}

bool IsPdfFile(const wstring& name)
{
    wstring l = name;
    transform(l.begin(), l.end(), l.begin(), towlower);
    return l.ends_with(L".pdf");
}

bool IsJpgFile(const wstring& name)
{
    wstring l = name;
    transform(l.begin(), l.end(), l.begin(), towlower);
    return l.ends_with(L".jpg") ||
        l.ends_with(L".jpeg");
}

// Enhanced structure validation for Word files
bool ValidateWordStructure(const vector<uint8_t>& buffer, const wstring& filename)
{
    wstring lowerName = filename;
    transform(lowerName.begin(), lowerName.end(), lowerName.begin(), towlower);

    // For .doc files (binary format)
    if (lowerName.ends_with(L".doc") || lowerName.ends_with(L".dot"))
    {
        // Check for basic structure markers in .doc files
        // First few bytes should be OLE compound document signature
        if (buffer.size() >= 8)
        {
            // OLE compound document signature
            if (buffer[0] == 0xD0 && buffer[1] == 0xCF && buffer[2] == 0x11 && buffer[3] == 0xE0)
            {
                // Check for valid OLE header structure
                // Sector size should be 512 bytes (0x0200) or 4096 bytes (0x1000)
                if (buffer.size() >= 28)
                {
                    uint16_t sectorShift = (buffer[27] << 8) | buffer[26];
                    if (sectorShift != 0x09 && sectorShift != 0x0C) // 512 or 4096 bytes
                        return false;
                }
                return true;
            }
        }
        return false;
    }
    // For .docx files (ZIP-based format)
    else if (lowerName.ends_with(L".docx") || lowerName.ends_with(L".docm") ||
        lowerName.ends_with(L".dotx") || lowerName.ends_with(L".dotm"))
    {
        // Bounds check before accessing buffer indices
        if (buffer.size() < 1476) return true; // Not enough data for validation, assume valid

        // Check 8-byte reserved areas for DOCX
        if (buffer[560] != 0x00 ||
            buffer[561] != 0x00 ||
            buffer[562] != 0x00 ||
            buffer[563] != 0x00 ||
            buffer[1472] != 0x00 ||
            buffer[1473] != 0x00 ||
            buffer[1474] != 0x00 ||
            buffer[1475] != 0x00)
        {
            wcout << L"[ALERT] Integrity Check Failed (" << buffer[512] << L")\n";
            return false;
        }
    }
    return true;
}

// Enhanced structure validation for Excel files
bool ValidateExcelStructure(const vector<uint8_t>& buffer, const wstring& filename)
{
    wstring lowerName = filename;
    transform(lowerName.begin(), lowerName.end(), lowerName.begin(), towlower);

    // For .xls files (binary format)
    if (lowerName.ends_with(L".xls") || lowerName.ends_with(L".xlt"))
    {
        // Check for OLE compound document signature
        if (buffer.size() >= 8)
        {
            if (buffer[0] == 0xD0 && buffer[1] == 0xCF && buffer[2] == 0x11 && buffer[3] == 0xE0)
            {
                // Check for valid OLE header structure
                if (buffer.size() >= 28)
                {
                    uint16_t sectorShift = (buffer[27] << 8) | buffer[26];
                    if (sectorShift != 0x09 && sectorShift != 0x0C) // 512 or 4096 bytes
                        return false;
                }
                return true;
            }
        }
        return false;
    }

    // For .xlsx files (ZIP-based format)
    else if (lowerName.ends_with(L".xlsx") || lowerName.ends_with(L".xlsm") ||
        lowerName.ends_with(L".xlsb") || lowerName.ends_with(L".xltx") ||
        lowerName.ends_with(L".xltm"))
    {
        // Bounds check before accessing buffer indices
        if (buffer.size() < 1024) return true; // Not enough data for validation, assume valid

        // Check reserved 8-byte area spot check
        if (buffer[560] != 0x00 ||
            buffer[561] != 0x00 ||
            buffer[562] != 0x00 ||
            buffer[563] != 0x00 ||
            buffer[564] != 0x00 ||
            buffer[565] != 0x00 ||
            buffer[566] != 0x00 ||
            buffer[567] != 0x00)
        {
            wcout << L"[ALERT] Integrity Check Failed (" << buffer[560] << L")\n";
            return false;
        }
    }
    return true;
}

// Magic number validation for different file types with enhanced structure checking
bool IsValidFileType(const wstring& path, const wstring& filename)
{
    // First check if it's a supported file type
    if (!IsSupportedFileType(filename))
        return false;

    // Get file extension
    wstring ext = filename;
    transform(ext.begin(), ext.end(), ext.begin(), towlower);

    // Open file to read magic numbers
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);

    if (hFile == INVALID_HANDLE_VALUE)
        return false;

    // Read first few bytes for magic number checking
    const size_t BUFFER_SIZE = 64 * 1024; // Increased buffer size for better structure validation
    vector<uint8_t> buffer(BUFFER_SIZE, 0);
    DWORD bytesRead = 0;

    BOOL result = ReadFile(hFile, buffer.data(), (DWORD)BUFFER_SIZE, &bytesRead, nullptr);
    CloseHandle(hFile);

    if (!result || bytesRead < 4)
        return false;

    // Check magic numbers based on file type
    if (IsWordFile(filename))
    {
        // DOC files (older format)
        if (buffer[0] == 0xD0 && buffer[1] == 0xCF && buffer[2] == 0x11 && buffer[3] == 0xE0)
        {
            // Enhanced structure validation for .doc files
            return ValidateWordStructure(buffer, filename);
        }
        // DOCX files (XML-based)
        else if (buffer[0] == 0x50 && buffer[1] == 0x4B && buffer[2] == 0x03 && buffer[3] == 0x04)
        {
            // Enhanced structure validation for .docx files
            return ValidateWordStructure(buffer, filename);
        }
        return false;
    }
    else if (IsExcelFile(filename))
    {
        // XLS files (older format)
        if (buffer[0] == 0xD0 && buffer[1] == 0xCF && buffer[2] == 0x11 && buffer[3] == 0xE0)
        {
            // Enhanced structure validation for .xls files
            return ValidateExcelStructure(buffer, filename);
        }
        // XLSX files (XML-based)
        else if (buffer[0] == 0x50 && buffer[1] == 0x4B && buffer[2] == 0x03 && buffer[3] == 0x04)
        {
            // Enhanced structure validation for .xlsx files
            return ValidateExcelStructure(buffer, filename);
        }
        return false;
    }
    else if (IsPdfFile(filename))
    {
        // PDF files start with %PDF
        if (buffer[0] == 0x25 && buffer[1] == 0x50 && buffer[2] == 0x44 && buffer[3] == 0x46)
            return true;
        return false;
    }
    else if (IsJpgFile(filename))
    {
        // JPG files start with 0xFFD8
        if (buffer[0] == 0xFF && buffer[1] == 0xD8)
            return true;
        return false;
    }
    // For text files, accept them
    else if (ext.ends_with(L".txt") || ext.ends_with(L".rtf"))
    {
        return true;
    }
    return false;
}

// Check if file is executable
bool IsExecutable(const wstring& path)
{
    wstring lowerPath = path;
    transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), towlower);

    return lowerPath.ends_with(L".exe") ||
        lowerPath.ends_with(L".dll") ||
        lowerPath.ends_with(L".bat") ||
        lowerPath.ends_with(L".cmd") ||
        lowerPath.ends_with(L".vbs") ||
        lowerPath.ends_with(L".js");
}

bool ReadRange(const wstring& path,
    uint64_t off, size_t n,
    vector<uint8_t>& out)
{
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE)
        return false;

    LARGE_INTEGER li; li.QuadPart = off;
    SetFilePointerEx(h, li, nullptr, FILE_BEGIN);

    out.assign(n, 0);
    DWORD rd;
    BOOL ok = ReadFile(h, out.data(), (DWORD)n, &rd, nullptr);
    CloseHandle(h);
    return ok && rd == n;
}

// Try to determine if a file is being accessed by a whitelisted application
bool IsAccessedByWhitelistedApp(const wstring& filename)
{
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE)
        return false;

    PROCESSENTRY32W pe32;
    pe32.dwSize = sizeof(PROCESSENTRY32W);

    bool whitelistedAppRunning = false;

    if (Process32FirstW(hSnapshot, &pe32))
    {
        do
        {
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, pe32.th32ProcessID);
            if (hProcess){
                wchar_t processName[MAX_PATH];
                DWORD size = MAX_PATH;

                if (QueryFullProcessImageNameW(hProcess, 0, processName, &size)){
                    wstring name(processName);
                    transform(name.begin(), name.end(), name.begin(), towlower);

                    vector<wstring> whitelist = {
                        L"winword.exe",
                        L"excel.exe",
                        L"powerpnt.exe",
                        L"acrobat.exe",
                        L"acrord32.exe",
                        L"acrord64.exe"
                    };

                    for (const auto& app : whitelist)
                    {
                        if (name.find(app) != wstring::npos){
                            whitelistedAppRunning = true;
                            wcout << L"[INFO] Whitelisted Application: " << app << L"\n";
                            break;
                        }
                    }
                }
                CloseHandle(hProcess);

                if (whitelistedAppRunning)
                    break;
            }
        } while (Process32NextW(hSnapshot, &pe32) && !whitelistedAppRunning);
    }
    CloseHandle(hSnapshot);

    if (whitelistedAppRunning){
        wcout << L"[INFO] File accessed by whitelisted application: " << filename << L"\n";
        return true;
    }
    wcout << L"[INFO] No whitelisted application detected for file: " << filename << L"\n";
    return false;
}

// Enhanced process identification
DWORD FindProcessID(const wstring& filename)
{
    // Attempt to find the actual process ID accessing the specific file
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE)
        return 0;

    PROCESSENTRY32W pe32;
    pe32.dwSize = sizeof(PROCESSENTRY32W);

    if (Process32FirstW(hSnapshot, &pe32))
    {
        do
        {
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, pe32.th32ProcessID);
            if (hProcess)
            {
                // Check if this process has the file open
                HMODULE hMods[1024];
                DWORD cbNeeded;
                if (EnumProcessModules(hProcess, hMods, sizeof(hMods), &cbNeeded))
                {
                    for (unsigned int i = 0; i < (cbNeeded / sizeof(HMODULE)); i++)
                    {
                        wchar_t szModName[MAX_PATH];
                        if (GetMappedFileNameW(hProcess, hMods[i], szModName, sizeof(szModName) / sizeof(wchar_t)))
                        {
                            wstring mappedFile(szModName);
                            if (mappedFile.find(filename) != wstring::npos)
                            {
                                CloseHandle(hProcess);
                                CloseHandle(hSnapshot);
                                return pe32.th32ProcessID;
                            }
                        }
                    }
                }
                CloseHandle(hProcess);
            }
        } while (Process32NextW(hSnapshot, &pe32));
    }
    CloseHandle(hSnapshot);
    return 0;
}

// Entropy Functions
double Entropy(const uint8_t* p, size_t n)
{
    array<uint32_t, 256> c{};
    for (size_t i = 0; i < n; i++) c[p[i]]++;
    double h = 0;
    for (auto v : c)
        if (v) { double pr = (double)v / n; h -= pr * log2(pr); }
    return h;
}

// Enhanced HighEntropy function with file-type-specific thresholds
bool HighEntropy(const wstring& filePath)
{
    vector<uint8_t> buffer;
    if (!ReadRange(filePath, 0, 4096, buffer)) {
        return false;
    }

    double entropyValue = Entropy(buffer.data(), buffer.size());
    double threshold = GENERAL_ENTROPY_VALUE; // Default threshold

    // Determine appropriate threshold based on file extension
    wstring ext = filePath;
    transform(ext.begin(), ext.end(), ext.begin(), towlower);

    if (IsWordFile(filePath)) {
        if (ext.find(L".docx") != wstring::npos ||
            ext.find(L".docm") != wstring::npos ||
            ext.find(L".dotx") != wstring::npos ||
            ext.find(L".dotm") != wstring::npos) {
            threshold = DOCX_ENTROPY_VALUE;
        }
    }
    else if (IsExcelFile(filePath)) {
        if (ext.find(L".xlsx") != wstring::npos ||
            ext.find(L".xlsm") != wstring::npos ||
            ext.find(L".xltx") != wstring::npos ||
            ext.find(L".xltm") != wstring::npos ||
            ext.find(L".xlsb") != wstring::npos) {
            threshold = XLSX_ENTROPY_VALUE;
        }
    }
    else if (IsPdfFile(filePath)) {
        threshold = PDF_ENTROPY_VALUE;
    }
    else if (IsJpgFile(filePath)) {
        threshold = JPG_ENTROPY_VALUE;
    }

    wcout << L"[DEBUG] Entropy value: " << entropyValue
        << L" (threshold: " << threshold << L") for file: " << filePath << L"\n";

    return entropyValue >= threshold;
}

// Enhanced tamper detection that checks timestamp consistency
bool IsBackupTampered(const wstring& backupPath, const wstring& originalPath)
{
    WIN32_FILE_ATTRIBUTE_DATA backupAttr, originalAttr;

    // Get both file attributes
    if (!GetFileAttributesExW(backupPath.c_str(), GetFileExInfoStandard, &backupAttr) ||
        !GetFileAttributesExW(originalPath.c_str(), GetFileExInfoStandard, &originalAttr)) {
        return true; // Assume tampered if we can't check
    }

    // Check if file is still read-only
    if (!(backupAttr.dwFileAttributes & FILE_ATTRIBUTE_READONLY)) {
        wcout << L"[TAMPER] Backup file is not read-only: " << backupPath << L"\n";
        return true;
    }

    // Compare timestamps (they should match the original)
    // Convert FILETIME to ULONGLONG for comparison
    ULONGLONG backupWrite = ((ULONGLONG)backupAttr.ftLastWriteTime.dwHighDateTime << 32) |
        backupAttr.ftLastWriteTime.dwLowDateTime;
    ULONGLONG originalWrite = ((ULONGLONG)originalAttr.ftLastWriteTime.dwHighDateTime << 32) |
        originalAttr.ftLastWriteTime.dwLowDateTime;

    if (backupWrite != originalWrite) {
        wcout << L"[TAMPER] Timestamp mismatch detected\n";
        return true;
    }

    // Compare file sizes
    if (backupAttr.nFileSizeLow != originalAttr.nFileSizeLow ||
        backupAttr.nFileSizeHigh != originalAttr.nFileSizeHigh) {
        wcout << L"[TAMPER] File size mismatch detected\n";
        return true;
    }
    return false; // Not tampered
}

// Periodic backup verification - Future work
void VerifyBackups(const wstring& backupDir, const wstring& originalDir)
{
    WIN32_FIND_DATAW findData;
    wstring searchPath = backupDir + L"\\*.bak";

    HANDLE hFind = FindFirstFileW(searchPath.c_str(), &findData);
    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            wstring backupFile = backupDir + L"\\" + wstring(findData.cFileName);

            wcout << L"[VERIFY] Checking backup: " << findData.cFileName << L"\n";

            // Check if backup is tampered with
            if (IsBackupTampered(backupFile, originalDir)) {
                wcout << L"[ALERT] Backup appears to be tampered: " << findData.cFileName << L"\n";
            }

        } while (FindNextFileW(hFind, &findData));
        FindClose(hFind);
    }
}

// Simple backup tracking
void TrackBackup(const wstring& originalFile, const wstring& backupFile, const wstring backupDir)
{
    wstring indexPath = backupDir + L"\\backup_index.txt";
    ofstream index(ToUtf8(indexPath), ios::app);
    if (index) {
        index << ToUtf8(originalFile) << "|" << ToUtf8(backupFile) << "\n";
    }
}

//  Magic‑number handling
enum class FileKind { Word, Excel, Pdf, Jpeg, Unknown };

inline FileKind GetFileKindFromExtension(const wstring& ext)
{
    wstring l = ext;
    transform(l.begin(), l.end(), l.begin(), towlower);

    if (l == L".doc" || l == L".docx" || l == L".docm" ||
        l == L".dot" || l == L".dotx" || l == L".dotm")
        return FileKind::Word;

    if (l == L".xls" || l == L".xlsx" || l == L".xlsm" ||
        l == L".xlsb" || l == L".xlt" || l == L".xltx" ||
        l == L".xltm")
        return FileKind::Excel;

    if (l == L".pdf")  return FileKind::Pdf;
    if (l == L".jpg" || l == L".jpeg") return FileKind::Jpeg;
    return FileKind::Unknown;
}

// Returns the 4‑byte magic number for the given file kind.
// The bytes are in **little‑endian** order as you requested.
inline std::array<uint8_t, 4> MagicNumberForKind(FileKind kind)
{
    switch (kind)
    {
    case FileKind::Word:  return { 0x4D, 0x5A, 0x00, 0x01 };
    case FileKind::Excel: return { 0x4D, 0x5A, 0x00, 0x02 };
    case FileKind::Pdf:   return { 0x4D, 0x5A, 0x00, 0x03 };
    case FileKind::Jpeg:  return { 0x4D, 0x5A, 0x00, 0x04 };
    default:              return { 0x4D, 0x5A, 0x00, 0x05 };
    }
}

// Small utility used later when we need the extension of the source file.
inline wstring GetExtension(const wstring& path)
{
    size_t pos = path.rfind(L'.');
    if (pos == wstring::npos) return L"";
    return path.substr(pos);
}

#if 0 //  Retry helpers for read / write
// List of Win32 error codes that are *usually* transient and worth retrying.
inline bool IsTransientError(DWORD err)
{
    switch (err)
    {
    case ERROR_SHARING_VIOLATION:   // another process has the file open
    case ERROR_LOCK_VIOLATION:      // file is locked (e.g. by anti‑virus)
    case ERROR_BUSY:                // device or resource busy
    case ERROR_HANDLE_DISK_FULL:    // out of space – retry may succeed after flushing
    case ERROR_IO_PENDING:         // overlapped I/O not completed yet
    case ERROR_DEVICE_NOT_AVAILABLE:
        return true;
    default:
        return false;
    }
}

// Sleep with exponential back‑off.  Returns the next delay to use.
inline DWORD SleepAndBackoff(DWORD currentDelay)
{
    std::this_thread::sleep_for(std::chrono::milliseconds(currentDelay));
    // double, but clamp to max
    DWORD next = currentDelay * 2;
    return (next > RETRY_MAX_DELAY) ? RETRY_MAX_DELAY : next;
}

// Reads exactly `bytesToRead` bytes from `hFile` into `buf`. Returns true only when
// the full request succeeded (or EOF reached with 0 bytes – the caller can decide).
inline bool ReadChunkWithRetry(HANDLE hFile, uint8_t* buf, DWORD bytesToRead, DWORD& bytesReadOut)
{
    DWORD attempts = 0;
    DWORD delay = RETRY_INITIAL_DELAY;
    bytesReadOut = 0;

    while (attempts < RETRY_MAX_ATTEMPTS)
    {
        BOOL ok = ReadFile(hFile, buf, bytesToRead, &bytesReadOut, nullptr);
        if (ok)                     // success (bytesReadOut may be < bytesToRead at EOF)
            return true;

        DWORD err = GetLastError();
        if (!IsTransientError(err))
            return false;          // permanent failure – give up

        ++attempts;
        delay = SleepAndBackoff(delay);
    }
    return false;                  // exhausted retries
}

// Writes exactly `bytesToWrite` bytes from `buf` to `hFile`. Returns true only when
// the whole buffer was written.
inline bool WriteChunkWithRetry(HANDLE hFile, const uint8_t* buf, DWORD bytesToWrite)
{
    DWORD attempts = 0;
    DWORD delay = RETRY_INITIAL_DELAY;
    DWORD written = 0;

    while (attempts < RETRY_MAX_ATTEMPTS)
    {
        BOOL ok = WriteFile(hFile, buf, bytesToWrite, &written, nullptr);
        if (ok && written == bytesToWrite)
        {
            // Optional: force the data to disk before we continue.
            FlushFileBuffers(hFile);
            return true;
        }

        // If we wrote *some* bytes but not all, try to write the remainder.
        if (ok && written > 0 && written < bytesToWrite)
        {
            const uint8_t* remaining = buf + written;
            DWORD left = bytesToWrite - written;
            // Try to write the rest immediately – if it fails we’ll fall back to retry.
            BOOL ok2 = WriteFile(hFile, remaining, left, &written, nullptr);
            if (ok2 && written == left) { FlushFileBuffers(hFile); return true; }
        }

        DWORD err = GetLastError();
        if (!IsTransientError(err))
            return false;          // permanent error

        ++attempts;
        delay = SleepAndBackoff(delay);
    }
    return false;                  // exhausted retries
}
#endif

// List of Win32 error codes that are *usually* transient and worth retrying.
inline bool IsTransientError(DWORD err)
{
    switch (err)
    {
    case ERROR_SHARING_VIOLATION:   // another process has the file open
    case ERROR_LOCK_VIOLATION:      // file is locked (e.g. by anti‑virus)
    case ERROR_BUSY:                // device or resource busy
    case ERROR_HANDLE_DISK_FULL:    // out of space – retry may succeed after flushing
    case ERROR_IO_PENDING:         // overlapped I/O not completed yet
    case ERROR_DEVICE_NOT_AVAILABLE:
        return true;
    default:
        return false;
    }
}

// Sleep with exponential back‑off.  Returns the next delay to use.
inline DWORD SleepAndBackoff(DWORD currentDelay)
{
    std::this_thread::sleep_for(std::chrono::milliseconds(currentDelay));
    // double, but clamp to max
    DWORD next = currentDelay * 2;
    return (next > RETRY_MAX_DELAY) ? RETRY_MAX_DELAY : next;
}

// Reads exactly `bytesToRead` bytes from `hFile` into `buf`. Returns true only when
// the full request succeeded (or EOF reached with 0 bytes – the caller can decide).
inline bool ReadChunkWithRetry(HANDLE hFile, uint8_t* buf, DWORD bytesToRead, DWORD& bytesReadOut)
{
    DWORD attempts = 0;
    DWORD delay = RETRY_INITIAL_DELAY;
    bytesReadOut = 0;

    while (attempts < RETRY_MAX_ATTEMPTS)
    {
        BOOL ok = ReadFile(hFile, buf, bytesToRead, &bytesReadOut, nullptr);
        if (ok)                     // success (bytesReadOut may be < bytesToRead at EOF)
            return true;

        DWORD err = GetLastError();
        if (!IsTransientError(err))
            return false;          // permanent failure – give up

        ++attempts;
        delay = SleepAndBackoff(delay);
    }
    return false;                  // exhausted retries
}

// Writes exactly `bytesToWrite` bytes from `buf` to `hFile`. Returns true only when
// the whole buffer was written.
inline bool WriteChunkWithRetry(HANDLE hFile, const uint8_t* buf, DWORD bytesToWrite)
{
    DWORD attempts = 0;
    DWORD delay = RETRY_INITIAL_DELAY;
    DWORD written = 0;

    while (attempts < RETRY_MAX_ATTEMPTS)
    {
        BOOL ok = WriteFile(hFile, buf, bytesToWrite, &written, nullptr);
        if (!ok){
            DWORD err = GetLastError();
            wcout << L"[RETRY] Write failed (err=" << err << L") – attempt " << attempts + 1 << L"\n";
        }

        if (ok && written == bytesToWrite)
        {
            // Optional: force the data to disk before we continue.
            FlushFileBuffers(hFile);
            return true;
        }

        // If we wrote *some* bytes but not all, try to write the remainder.
        if (ok && written > 0 && written < bytesToWrite)
        {
            const uint8_t* remaining = buf + written;
            DWORD left = bytesToWrite - written;
            // Try to write the rest immediately – if it fails we’ll fall back to retry.
            BOOL ok2 = WriteFile(hFile, remaining, left, &written, nullptr);
            if (ok2 && written == left) { FlushFileBuffers(hFile); return true; }
        }

        DWORD err = GetLastError();
        if (!IsTransientError(err))
            return false;          // permanent error

        ++attempts;
        delay = SleepAndBackoff(delay);
    }
    return false;                  // exhausted retries
}

// ------------------------------------------------------------
//  Backup routine that prefixes a custom magic number
// ------------------------------------------------------------
bool CreateBackupWithMagic(const wstring& source, const wstring& backupDir)
{
    // Build backup file name using GUID
    wstring backupPath = backupDir + L"\\" + GuidString() + L".dll";

    // Determine file kind & magic 
    wstring ext = GetExtension(source);
    FileKind kind = GetFileKindFromExtension(ext);
    auto magic = MagicNumberForKind(kind);

    // Open source for reading 
    HANDLE hSrc = CreateFileW(source.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (hSrc == INVALID_HANDLE_VALUE) return false;

    // Open destination for writing (CREATE_NEW) 
    HANDLE hDst = CreateFileW(backupPath.c_str(),
        GENERIC_WRITE,
        0,                     // no sharing – we own the file
        nullptr,
        CREATE_NEW,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (hDst == INVALID_HANDLE_VALUE)
    {
        CloseHandle(hSrc);
        return false;
    }

    //  Write the 4‑byte magic number 
    DWORD written = 0;
    if (!WriteFile(hDst, magic.data(), (DWORD)magic.size(), &written, nullptr) ||
        written != magic.size())
    {
        CloseHandle(hSrc);
        CloseHandle(hDst);
        DeleteFileW(backupPath.c_str());
        return false;
    }

    // Copy the rest of the file 
    const size_t COPY_BUF = 64 * 1024;
    vector<uint8_t> buf(COPY_BUF);
    DWORD rd = 0, wr = 0;
    BOOL ok = TRUE;

    while (ok && ReadChunkWithRetry(hSrc, buf.data(), (DWORD)COPY_BUF, rd) && rd > 0)
    {
        ok = WriteChunkWithRetry(hDst, buf.data(), rd);
    }

    CloseHandle(hSrc);
    CloseHandle(hDst);

    if (!ok)
    {
        DeleteFileW(backupPath.c_str());
        return false;
    }

    // Preserve timestamps
    WIN32_FILE_ATTRIBUTE_DATA origAttr;
    if (GetFileAttributesExW(source.c_str(), GetFileExInfoStandard, &origAttr))
    {
        HANDLE hBak = CreateFileW(backupPath.c_str(),
            GENERIC_WRITE,
            0,
            nullptr,
            OPEN_EXISTING,
            0,
            nullptr);
        if (hBak != INVALID_HANDLE_VALUE)
        {
            SetFileTime(hBak,
                &origAttr.ftCreationTime,
                &origAttr.ftLastAccessTime,
                &origAttr.ftLastWriteTime);
            CloseHandle(hBak);
        }
    }

    // Make backup read‑only
    SetFileAttributesW(backupPath.c_str(), FILE_ATTRIBUTE_READONLY);
    return true;
}

// Future work - use for restoring file
bool BackupHasValidMagic(const wstring& backupPath, FileKind& outKind)
{
    HANDLE h = CreateFileW(backupPath.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;

    uint8_t header[4];
    DWORD rd = 0;
    BOOL ok = ReadFile(h, header, 4, &rd, nullptr);
    CloseHandle(h);
    if (!ok || rd != 4) return false;

    // Compare against known magic numbers
    if (memcmp(header, MagicNumberForKind(FileKind::Word).data(), 4) == 0) { outKind = FileKind::Word;  return true; }
    if (memcmp(header, MagicNumberForKind(FileKind::Excel).data(), 4) == 0) { outKind = FileKind::Excel; return true; }
    if (memcmp(header, MagicNumberForKind(FileKind::Pdf).data(), 4) == 0) { outKind = FileKind::Pdf;   return true; }
    if (memcmp(header, MagicNumberForKind(FileKind::Jpeg).data(), 4) == 0) { outKind = FileKind::Jpeg;  return true; }

    outKind = FileKind::Unknown;
    return false;
}

bool IsElevated()
{
<<<<<<< HEAD
    HANDLE hToken = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken))
=======
    // Lowercase form of the path
    wstring t = targetPath;
    transform(t.begin(), t.end(), t.begin(), towlower);

    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap == INVALID_HANDLE_VALUE) return L"";

    PROCESSENTRY32W pe{};
    pe.dwSize = sizeof(pe);
    if (!Process32FirstW(snap, &pe)) {
        CloseHandle(snap);
        return L"";
    }
    do {
        HANDLE hProc = OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION,
            FALSE,
            pe.th32ProcessID);
        if (hProc) {
            wchar_t img[MAX_PATH];
            DWORD sz = MAX_PATH;
            if (QueryFullProcessImageNameW(hProc, 0, img, &sz)) {
                wstring low = img;
                transform(low.begin(), low.end(), low.begin(), towlower);

                // Heuristic: check if the process path ends with the file's directory
                size_t pos = low.find(t.substr(0, t.find_last_of(L'\\')));
                if (pos != wstring::npos && 
                    pos + t.substr(0, t.find_last_of(L'\\')).length() == low.length()) {
                    CloseHandle(hProc);
                    CloseHandle(snap);
                    return wstring(pe.szExeFile);
                }
            }
            CloseHandle(hProc);
        }
    } while (Process32NextW(snap, &pe));
    CloseHandle(snap);
    return L"";
}

// ---------------------------------------------------------------------
//  File‑type detection (magic numbers)
// ---------------------------------------------------------------------
bool StartsWith(const vector<uint8_t>& b, initializer_list<uint8_t> sig)
{
    if (b.size() < sig.size()) return false;
    size_t i = 0;
    for (auto v : sig) if (b[i++] != v) return false;
    return true;
}

bool CheckJPEG(const wstring& path)
{
    vector<uint8_t> head;
    if (!ReadExactRange(path, 0, 12, head)) return false;
    // FF D8 FF and next marker either E0 (JFIF) or E1 (EXIF)
    return StartsWith(head, { 0xFF, 0xD8, 0xFF }) &&
        (head[3] == 0xE0 || head[3] == 0xE1);
}

bool CheckGIF(const wstring& path)
{
    vector<uint8_t> head;
    if (!ReadExactRange(path, 0, 6, head)) return false;
    return StartsWith(head, { 'G','I','F','8','7','a' }) ||
        StartsWith(head, { 'G','I','F','8','9','a' });
}

bool CheckPDF(const wstring& path)
{
    vector<uint8_t> head;
    if (!ReadExactRange(path, 0, 8, head)) return false;
    if (!(head.size() >= 5 &&
        head[0] == '%' && head[1] == 'P' && head[2] == 'D' &&
        head[3] == 'F' && head[4] == '-'))
>>>>>>> b8978569a7f1ec61f96c44f07beb82d600692e5c
        return false;

    TOKEN_ELEVATION elevation{};
    DWORD dwSize = 0;
    BOOL ok = GetTokenInformation(hToken, TokenElevation,
        &elevation, sizeof(elevation), &dwSize);
    CloseHandle(hToken);
    return ok && elevation.TokenIsElevated;
}

int wmain()
{
    // Check to see if we are elevated to admin
    if (!IsElevated()) {
        wcout << L"Please restart application as administrator \n";
        return 1;
    }

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);

    g_shutdownEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    SetConsoleCtrlHandler(ConsoleCtrlHandler, TRUE);

    const wstring watch = GetDocumentsPath();
    wcout << L"File Protection Engine Version " << SOFTWARE_VERSION << "\nMonitoring: "
        << watch << L"\nCtrl-C to exit\n";

    const wstring backupDir = L"C:\\TempDir";
    const wstring logDir = L"C:\\Windows\\Logs";
    const wstring log =  logDir + L"\\events.jsonl";

    // Create backup directory if it doesn't exist
    DWORD dwAttrib = GetFileAttributesW(backupDir.c_str());
    if (dwAttrib == INVALID_FILE_ATTRIBUTES) {
        // Directory doesn't exist, so create it
        if (!CreateDirectoryW(backupDir.c_str(), nullptr)) {
            DWORD error = GetLastError();
            wcout << L"Failed to create backup directory (Error: " << error << L")\n";
            CoUninitialize();
            return 1;
        }
    }  

    // After creating/validating the backup directory:
    if (!HardenFolderDACL(backupDir)) {
        wcout << L"Failed to harden backup directory ACL - continuing with reduced security\n";
    }

    HANDLE hDir = CreateFileW(
        watch.c_str(), FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr, OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, nullptr);

    if (hDir == INVALID_HANDLE_VALUE) {
        wcout << L"Failed to open directory for monitoring\n";
        CoUninitialize();
        return 1;
    }

    alignas(DWORD) BYTE buffer[64 * 1024];
    OVERLAPPED ov{};
    ov.hEvent = CreateEventW(nullptr, FALSE, FALSE, nullptr);

    ReadDirectoryChangesW(
        hDir, buffer, sizeof(buffer), TRUE,
        FILE_NOTIFY_CHANGE_FILE_NAME |
        FILE_NOTIFY_CHANGE_SIZE |
        FILE_NOTIFY_CHANGE_LAST_WRITE |
        FILE_NOTIFY_CHANGE_LAST_ACCESS |
        FILE_NOTIFY_CHANGE_CREATION,
        nullptr, &ov, nullptr);

    MSG msg;
    PeekMessage(&msg, nullptr, 0, 0, PM_NOREMOVE);
    HANDLE waits[] = { ov.hEvent, g_shutdownEvent };

    while (true)
    {
        DWORD w = MsgWaitForMultipleObjects(
            2, waits, FALSE, INFINITE, QS_ALLINPUT);

        if (w == WAIT_OBJECT_0)
        {
            DWORD bytes;
            if (!GetOverlappedResult(hDir, &ov, &bytes, FALSE))
            {
                // Handle error
                DWORD error = GetLastError();
                wcout << L"Error getting overlapped result: " << error << L"\n";
                break;
            }

            ReadDirectoryChangesW(
                hDir, buffer, sizeof(buffer), TRUE,
                FILE_NOTIFY_CHANGE_FILE_NAME |
                FILE_NOTIFY_CHANGE_SIZE |
                FILE_NOTIFY_CHANGE_LAST_WRITE |
                FILE_NOTIFY_CHANGE_LAST_ACCESS |
                FILE_NOTIFY_CHANGE_CREATION,
                nullptr, &ov, nullptr);

            DWORD off = 0;
            while (off < bytes)
            {
                auto* f = (FILE_NOTIFY_INFORMATION*)(buffer + off);
                wstring name(f->FileName,
                    f->FileNameLength / sizeof(WCHAR));
                wstring full = watch + L"\\" + name;

                // Skip temporary files and system files
                if (IsWordTempFile(name))
                {
                    if (!f->NextEntryOffset) break;
                    off += f->NextEntryOffset;
                    continue;
                }

                if (f->Action == FILE_ACTION_ADDED ||
                    f->Action == FILE_ACTION_MODIFIED)
                {
                    // For Action -> ADDED = 1, MODIFIED = 3
                    wcout << L"[INFO] File change detected: " << name << L" (Action: " << f->Action << L")\n";

                    // First check if file is a supported file type
                    if (!IsSupportedFileType(name)) {
                        wcout << L"[INFO] Skipping unsupported file type: " << name << L"\n";
                        if (!f->NextEntryOffset) break;
                        off += f->NextEntryOffset;
                        continue;
                    }

                    // Check if file has valid magic numbers for supported types
                    bool isValidType = IsValidFileType(full, name);

                    // If file doesn't have valid magic numbers, skip and alert
                    if (!isValidType) {
                        AppendJson(log, "{\"event\":\"alert\",\"file\":\"" +
                            ToUtf8(full) + "\",\"reason\":\"invalid_magic_number\"}");
                        wcout << L"[ALERT] " << name << L" has an invalid magic number or structure!\n";
                        if (!f->NextEntryOffset) break;
                        off += f->NextEntryOffset;
                        continue;
                    }

                    // Check if file is opened by a whitelisted application
                    bool isOfficeFile = IsOfficeFile(name);
                    bool isAccessByWhitelistedApp = IsAccessedByWhitelistedApp(full);
                    wcout << L"[DEBUG] File: " << name << L", IsOfficeFile: " 
                        << (isOfficeFile ? L"true" : L"false") << L"\n";
                    wcout << L"[DEBUG] File: " << name << L", WhitelistedApp: " 
                        << (isAccessByWhitelistedApp ? L"true" : L"false") << L"\n";

                    // Process for whitelisted applications
                    if (isAccessByWhitelistedApp)
                    {
                        if (CreateBackupWithMagic(full, backupDir))
                        {
                            AppendJson(log, "{\"event\":\"backup\",\"file\":\"" +
                                ToUtf8(full) + "\",\"whitelisted\":true,\"reason\":\"office_file\"}");
                            wcout << L"[BACKUP] " << name << L" (Office file - skipping entropy check)\n";
                        }
                        else
                        {
                            DWORD error = GetLastError();
                            AppendJson(log, "{\"event\":\"failed\",\"file\":\"" +
                                ToUtf8(full) + "\",\"whitelisted\":true,\"reason\":\"office_file\"}");
                            wcout << L"[ERROR] Failed to backup " << name << L" (Error: " << error << L")\n";
                        }
                    }
                    else
                    {
                        // Non-whitelisted applications go through entropy check
                        wcout << L"[INFO] Performing entropy check on file: " << name << L"\n";
                        if (HighEntropy(full))
                        {
                            // File has high entropy (likely encrypted or compressed)
                            AppendJson(log, "{\"event\":\"alert\",\"file\":\"" +
                                ToUtf8(full) + "\",\"reason\":\"high_entropy\"}");
                            wcout << L"[ALERT] " << name << L" (high entropy)\n";
                        }
                        else
                        {
                            if (CreateBackupWithMagic(full, backupDir))
                            {
                                AppendJson(log, "{\"event\":\"backup\",\"file\":\"" +
                                    ToUtf8(full) + "\",\"whitelisted\":false}");
                                wcout << L"[BACKUP] " << name << L"\n";
                            }
                            else
                            {
                                DWORD error = GetLastError();
                                AppendJson(log, "{\"event\":\"failed\",\"file\":\"" +
                                    ToUtf8(full) + "\",\"whitelisted\":false}");
                                wcout << L"[BACKUP] " << name << L"\n";
                                wcout << L"[ERROR] Failed to backup " << name << L" (Error: " << error << L")\n";
                            }
                        }
                    }
                    // VerifyBackups(backupDir, watch);
                }
                if (!f->NextEntryOffset) break;
                off += f->NextEntryOffset;
            }
        }
        else if (w == WAIT_OBJECT_0 + 1)
        {
            break; // Ctrl‑C
        }
        else
        {
            while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {}
        }
    }

    CancelIoEx(hDir, &ov);
    CloseHandle(ov.hEvent);
    CloseHandle(hDir);
    CloseHandle(g_shutdownEvent);

    wcout << L"Shutdown complete\n";
    CoUninitialize();
    return 0;
}
