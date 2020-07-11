#include "framework.h"
#include "resource.h"
#include "AutoWaitCursor.h"
#include "MessageCrackers.h"
#include "XmlHelpers.h"
#include "WindowHelpers.h"
#include "TcxRouteNameDialog.h"
#include <winioctl.h>

TcxRouteNameDialog::TcxRouteNameDialog(HWND hWnd)
    : m_hWnd(hWnd)
    , m_fIgnoreDeviceChange(false)
    , m_fInputFileDeleted(false)
{
}

TcxRouteNameDialog::~TcxRouteNameDialog()
{
}

BOOL TcxRouteNameDialog::SetupDeviceArrivalNotifications(IN GUID InterfaceClassGuid, OUT HDEVNOTIFY* hDeviceNotify)
{
    DEV_BROADCAST_DEVICEINTERFACE NotificationFilter = {};
    NotificationFilter.dbcc_size = sizeof(DEV_BROADCAST_DEVICEINTERFACE);
    NotificationFilter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
    NotificationFilter.dbcc_classguid = InterfaceClassGuid;

    *hDeviceNotify = RegisterDeviceNotification(m_hWnd, &NotificationFilter, DEVICE_NOTIFY_WINDOW_HANDLE);
    if (NULL == *hDeviceNotify)
    {
        return FALSE;
    }

    return TRUE;
}


LRESULT TcxRouteNameDialog::OnInitDialog(WPARAM, LPARAM)
{
    // Limit the route name to 32 characters
    SendDlgItemMessageW(m_hWnd, IDC_EDIT_ROUTE_NAME, EM_SETLIMITTEXT, 32, 0);

    // Set window icons
    SendMessageW(m_hWnd, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(LoadImage(HINST_THIS_MODULE, MAKEINTRESOURCE(IDI_TCXROUTENAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR)));
    SendMessageW(m_hWnd, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(LoadImage(HINST_THIS_MODULE, MAKEINTRESOURCE(IDI_TCXROUTENAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR)));

    // Center the window on the desktop
    WindowHelpers::CenterWindow(m_hWnd, GetDesktopWindow());

    // If the application was provided any arguments, assume the first one is an input filename
    if (__argc > 1)
    {
        LoadFile(__wargv[1]);
    }

    // If the document is NULL, we were not able to successfully load a document
    if (m_spDocument == NULL)
    {
        // Set the focus to the browse button
        SetFocus(GetDlgItem(m_hWnd, IDC_BROWSE_INPUT));
    }
    else
    {
        // Set the focus to the route name edit control
        SetFocus(GetDlgItem(m_hWnd, IDC_EDIT_ROUTE_NAME));

        // Select all of the text
        SendDlgItemMessage(m_hWnd, IDC_EDIT_ROUTE_NAME, EM_SETSEL, 0, -1);
    }


    return 0;
}

void TcxRouteNameDialog::OnCancel(WPARAM, LPARAM)
{
    EndDialog(m_hWnd, IDCANCEL);
}

void TcxRouteNameDialog::ReplaceCharactersInPlace(PWSTR pszString, WCHAR chFrom, WCHAR chTo)
{
    for (PWSTR pszCurr = pszString; *pszCurr; ++pszCurr)
    {
        if (*pszCurr == chFrom)
        {
            *pszCurr = chTo;
        }
    }

}

void TcxRouteNameDialog::UpdateSaveButtonState()
{
    static int aEditControls[] = { IDC_EDIT_INPUT_FILE, IDC_EDIT_OUTPUT_FILE, IDC_EDIT_ROUTE_NAME };

    // Assume save will be enabled
    bool fEnableSave = true;
    for (int i = 0; i < ARRAYSIZE(aEditControls); ++i)
    {
        // If this edit control is empty, disable save
        if (GetWindowTextLengthW(GetDlgItem(m_hWnd, aEditControls[i])) == 0)
        {
            fEnableSave = false;
            break;
        }
    }

    // Is the save button currently enabled?
    bool fIsSaveEnabled = IsWindowEnabled(GetDlgItem(m_hWnd, IDC_SAVE));

    // If the old state and new state don't match, update the state
    if (fIsSaveEnabled != fEnableSave)
    {
        EnableWindow(GetDlgItem(m_hWnd, IDC_SAVE), fEnableSave);
    }
}

void TcxRouteNameDialog::UpdateDeleteButtonState()
{
    BOOL fEnable = TRUE;
    if (m_fInputFileDeleted)
    {
        fEnable = FALSE;
    }
    if (::GetWindowTextLengthW(GetDlgItem(m_hWnd, IDC_EDIT_INPUT_FILE)) == 0)
    {
        fEnable = FALSE;
    }
    if (::IsWindowEnabled(GetDlgItem(m_hWnd, IDC_DELETE_INPUT)) != fEnable)
    {
        ::EnableWindow(GetDlgItem(m_hWnd, IDC_DELETE_INPUT), fEnable);
    }
}

void TcxRouteNameDialog::OnEditChange(WPARAM, LPARAM)
{
    UpdateSaveButtonState();
}

int TcxRouteNameDialog::FormattedMessageBox(int nButtons, UINT uFormatStringId, ...)
{
    // Load the title string
    WCHAR szTitle[128] = {};
    ::LoadStringW(HINST_THIS_MODULE, IDS_APP_NAME, szTitle, ARRAYSIZE(szTitle));

    // Load the format string
    WCHAR szFormat[1024] = {};
    ::LoadStringW(HINST_THIS_MODULE, uFormatStringId, szFormat, ARRAYSIZE(szFormat));

    // Format the string
    WCHAR szMessage[1024] = {};
    va_list argList;
    va_start(argList, uFormatStringId);
    (void)::StringCchVPrintfW(szMessage, ARRAYSIZE(szMessage), szFormat, argList);
    va_end(argList);

    // Display the message
    return ::MessageBoxW(m_hWnd, szMessage, szTitle, nButtons);
}

void TcxRouteNameDialog::FormatStatusMessage(UINT uFormatStringId, ...)
{
    // Load the format string
    WCHAR szFormat[1024] = {};
    ::LoadStringW(HINST_THIS_MODULE, uFormatStringId, szFormat, ARRAYSIZE(szFormat));

    // Format the string
    WCHAR szMessage[1024] = {};
    va_list argList;
    va_start(argList, uFormatStringId);
    (void)::StringCchVPrintfW(szMessage, ARRAYSIZE(szMessage), szFormat, argList);
    va_end(argList);

    // Set the status text
    ::SetDlgItemText(m_hWnd, IDC_STATUS_BAR, szMessage);
}

HRESULT TcxRouteNameDialog::GetFirstGarminDeviceNewFilesDirectory(PWSTR pszDrive, size_t cchDrive)
{
    HRESULT hr = E_FAIL;
    DWORD dwDrives = GetLogicalDrives();
    for (int i = 0; i < 32; ++i)
    {
        WCHAR szRoot[MAX_PATH] = L"A:\\";
        if ((1 << i) & dwDrives)
        {
            szRoot[0] = 'A' + i;

            WCHAR szVolumeName[MAX_PATH];
            if (GetVolumeInformationW(szRoot, szVolumeName, ARRAYSIZE(szVolumeName), NULL, NULL, NULL, NULL, 0))
            {
                if (_wcsicmp(szVolumeName, L"GARMIN") == 0)
                {
                    WCHAR szNewFilesPath[MAX_PATH] = {};
                    if (SUCCEEDED(StringCchCopyW(szNewFilesPath, ARRAYSIZE(szNewFilesPath), szRoot)))
                    {
                        if (PathAppendW(szNewFilesPath, L"Garmin\\NewFiles"))
                        {
                            DWORD dwFileAttribues = GetFileAttributesW(szNewFilesPath);
                            if (dwFileAttribues != INVALID_FILE_ATTRIBUTES && (dwFileAttribues & FILE_ATTRIBUTE_DIRECTORY) != 0)
                            {
                                return StringCchCopyW(pszDrive, cchDrive, szNewFilesPath);
                            }
                        }
                    }
                }
            }
        }
    }
    return hr;
}

HRESULT TcxRouteNameDialog::LoadXmlDocument(PCWSTR pszFile, IXMLDOMDocument** ppXmlDocument)
{
    *ppXmlDocument = NULL;

    // Load the new document
    CComPtr<IXMLDOMDocument> spDocument;
    HRESULT hr = ::CoCreateInstance(__uuidof(DOMDocument), NULL, CLSCTX_ALL, IID_PPV_ARGS(&spDocument));
    if (SUCCEEDED(hr))
    {
        if (PathFileExistsW(pszFile))
        {
            VARIANT_BOOL fIsSuccessful = VARIANT_FALSE;
            hr = spDocument->load(CComVariant(pszFile), &fIsSuccessful);
            if (S_FALSE != hr && VARIANT_FALSE != fIsSuccessful)
            {
                *ppXmlDocument = spDocument.Detach();
            }
            else
            {
                // this isn't a valid XML file.
                hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        else
        {
            hr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
    }
    return hr;
}

HRESULT TcxRouteNameDialog::ValidateTcxRoute(IXMLDOMDocument* pXmlDocument)
{
    // Get the existing unique ID to verify this is a TCX route
    CComBSTR sbstrUniqueId;
    HRESULT hr = XmlHelpers::GetString(pXmlDocument, s_pszXPathIdQuery, &sbstrUniqueId);
    if (SUCCEEDED(hr))
    {
        // Get the course name to verify this is a TCX route
        CComBSTR sbstrName;
        hr = XmlHelpers::GetString(pXmlDocument, s_pszXPathNameQuery, &sbstrName);
        if (SUCCEEDED(hr))
        {
            // Create a new unique ID
            GUID guidUniqueId;
            hr = ::CoCreateGuid(&guidUniqueId);
            if (SUCCEEDED(hr))
            {
                // Format it
                WCHAR szGuid[MAX_PATH];
                hr = StringCchPrintfW(szGuid, ARRAYSIZE(szGuid), L"%08X%04X%04X%02X%02X%02X%02X%02X%02X%02X%02X", guidUniqueId.Data1, guidUniqueId.Data2, guidUniqueId.Data3, guidUniqueId.Data4[0], guidUniqueId.Data4[1], guidUniqueId.Data4[2], guidUniqueId.Data4[3], guidUniqueId.Data4[4], guidUniqueId.Data4[5], guidUniqueId.Data4[6], guidUniqueId.Data4[7]);
                if (SUCCEEDED(hr))
                {
                    // Set the new unique ID
                    hr = XmlHelpers::SetString(pXmlDocument, s_pszXPathIdQuery, szGuid);
                }
            }
        }
    }
    return hr;
}


HRESULT TcxRouteNameDialog::LoadFile(PCWSTR pszFile)
{
    AutoWaitCursor wc;

    // Load the new document
    CComPtr<IXMLDOMDocument> spDocument;
    HRESULT hr = LoadXmlDocument(pszFile, &spDocument);
    if (SUCCEEDED(hr))
    {
        // Verify it is a course and set its unique ID
        hr = ValidateTcxRoute(spDocument);
        if (SUCCEEDED(hr))
        {
            // Format the filename as a title
            WCHAR szFilename[MAX_PATH] = {};
            hr = StringCchCopyW(szFilename, ARRAYSIZE(szFilename), PathFindFileNameW(pszFile));
            if (SUCCEEDED(hr))
            {
                // Remove the extension
                PathRemoveExtensionW(szFilename);

                // Replace underscores with spaces
                ReplaceCharactersInPlace(szFilename, L'_', L' ');

                // Set the input file name
                SetDlgItemTextW(m_hWnd, IDC_EDIT_INPUT_FILE, pszFile);

                // Set the title
                SetDlgItemTextW(m_hWnd, IDC_EDIT_ROUTE_NAME, szFilename);

                // Empty the output file name (it will get filled in on the next step)
                SetDlgItemTextW(m_hWnd, IDC_EDIT_OUTPUT_FILE, L"");

                // Set the output file name if we have a device
                ConstructOutputFileName();

                // Set status bar text
                FormatStatusMessage(IDS_STATUS_FILE_LOADED, pszFile);

                // Reset the file deleted flag
                m_fInputFileDeleted = false;

                // Save the document pointer
                m_spDocument.Release();
                m_spDocument.Attach(spDocument.Detach());
            }
            else
            {
                FormattedMessageBox(MB_ICONEXCLAMATION | MB_OK, IDS_LOAD_ERROR_UNEXPECTED, pszFile, hr);
            }
        }
        else
        {
            FormattedMessageBox(MB_ICONEXCLAMATION | MB_OK, IDS_LOAD_ERROR_INVALID_TCX_ROUTE, pszFile, hr);
        }
    }
    else
    {
        FormattedMessageBox(MB_ICONEXCLAMATION | MB_OK, IDS_LOAD_ERROR_INVALID_FILE, pszFile, hr);
    }

    // Disable/Enable the Save button
    UpdateSaveButtonState();

    // Update the delete button state
    UpdateDeleteButtonState();
    return hr;
}

void TcxRouteNameDialog::OnBrowseInput(WPARAM, LPARAM)
{
    // Load the filter string
    HRESULT hr = S_OK;
    WCHAR szFilter[MAX_PATH] = {};
    if (LoadStringW(HINST_THIS_MODULE, IDS_OPEN_MASK, szFilter, ARRAYSIZE(szFilter)))
    {
        ReplaceCharactersInPlace(szFilter, L'|', L'\0');

        WCHAR szFileName[MAX_PATH] = {};

        // Show the dialog
        OPENFILENAME ofn = {};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = m_hWnd;
        ofn.lpstrFilter = szFilter;
        ofn.lpstrFile = szFileName;
        ofn.nMaxFile = ARRAYSIZE(szFileName);
        ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST;
        if (GetOpenFileNameW(&ofn))
        {
            LoadFile(szFileName);
        }
    }
    else
    {
        hr = HRESULT_FROM_WIN32(::GetLastError());
    }
}

void TcxRouteNameDialog::OnBrowseOutput(WPARAM, LPARAM)
{
    // Don't auto-update the name while we're looking for it
    m_fIgnoreDeviceChange = true;

    // Load the filter string
    HRESULT hr = S_OK;
    WCHAR szFilter[MAX_PATH] = {};
    if (LoadStringW(HINST_THIS_MODULE, IDS_OPEN_MASK, szFilter, ARRAYSIZE(szFilter)))
    {
        ReplaceCharactersInPlace(szFilter, L'|', L'\0');

        WCHAR szFileName[MAX_PATH] = {};
        if (GetDlgItemTextW(m_hWnd, IDC_EDIT_OUTPUT_FILE, szFileName, ARRAYSIZE(szFileName)) == 0)
        {
            WCHAR szInputFile[MAX_PATH] = {};
            if (GetDlgItemTextW(m_hWnd, IDC_EDIT_INPUT_FILE, szInputFile, ARRAYSIZE(szInputFile)) != 0)
            {
                StringCchCopyW(szFileName, ARRAYSIZE(szFileName), PathFindFileNameW(szInputFile));
            }
        }

        // Show the dialog
        OPENFILENAME ofn = {};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = m_hWnd;
        ofn.lpstrFilter = szFilter;
        ofn.lpstrFile = szFileName;
        ofn.nMaxFile = ARRAYSIZE(szFileName);
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
        ofn.lpstrDefExt = L"tcx";
        if (GetSaveFileNameW(&ofn))
        {
            SetDlgItemTextW(m_hWnd, IDC_EDIT_OUTPUT_FILE, szFileName);
        }
    }
    else
    {
        hr = HRESULT_FROM_WIN32(::GetLastError());
    }

    m_fIgnoreDeviceChange = false;

    // If the user cancelled, might as well try again.
    ConstructOutputFileName();
}

void TcxRouteNameDialog::OnDeleteInput(WPARAM, LPARAM)
{
    WCHAR szInputFileName[MAX_PATH] = {};
    if (GetDlgItemTextW(m_hWnd, IDC_EDIT_INPUT_FILE, szInputFileName, ARRAYSIZE(szInputFileName)))
    {
        // Make sure it isn't empty
        if (lstrlenW(szInputFileName) != 0)
        {
            if (::PathFileExistsW(szInputFileName))
            {
                BOOL fResult = ::DeleteFileW(szInputFileName);
                if (fResult)
                {
                    m_fInputFileDeleted = true;
                    FormatStatusMessage(IDS_STATUS_FILE_DELETED, szInputFileName);
                    FormattedMessageBox(MB_OK | MB_ICONASTERISK, IDS_STATUS_FILE_DELETED, szInputFileName);
                }
                else
                {
                    FormattedMessageBox(MB_OK | MB_ICONEXCLAMATION, IDS_ERROR_UNABLE_TO_DELETE, szInputFileName, HRESULT_FROM_WIN32(::GetLastError()));
                }
            }
            else
            {
                FormattedMessageBox(MB_OK | MB_ICONASTERISK, IDS_STATUS_INPUT_FILE_NOT_FOUND, szInputFileName);
            }
        }
    }
    UpdateDeleteButtonState();
}

LRESULT TcxRouteNameDialog::OnDeviceChange(WPARAM wParam, LPARAM lParam)
{
    if (!m_fIgnoreDeviceChange)
    {
        if (GetWindowTextLengthW(GetDlgItem(m_hWnd, IDC_EDIT_OUTPUT_FILE)) == 0)
        {
            if (wParam == DBT_DEVICEARRIVAL)
            {
                ConstructOutputFileName();
            }
        }
    }
    return TRUE;
}

void TcxRouteNameDialog::ConstructOutputFileName()
{
    // Get the input file name
    WCHAR szInputFileName[MAX_PATH] = {};
    if (GetDlgItemTextW(m_hWnd, IDC_EDIT_INPUT_FILE, szInputFileName, ARRAYSIZE(szInputFileName)))
    {
        // Make sure it isn't empty
        if (lstrlenW(szInputFileName) != 0)
        {
            // If the output file name IS empty
            if (GetWindowTextLengthW(GetDlgItem(m_hWnd, IDC_EDIT_OUTPUT_FILE)) == 0)
            {
                // Try to find the first Garmin device
                WCHAR szOutputFile[MAX_PATH] = { 0 };
                if (SUCCEEDED(GetFirstGarminDeviceNewFilesDirectory(szOutputFile, ARRAYSIZE(szOutputFile))))
                {
                    // Append the file name
                    PathAppendW(szOutputFile, PathFindFileNameW(szInputFileName));

                    // Set it
                    SetDlgItemTextW(m_hWnd, IDC_EDIT_OUTPUT_FILE, szOutputFile);
                }
            }
        }
    }
}


HRESULT TcxRouteNameDialog::SaveXmlDocumentAndFlush(PCWSTR pszFile)
{
    HRESULT hr = S_OK;

    HGLOBAL	hMem = ::GlobalAlloc(GMEM_MOVEABLE, 0);
    if (hMem != NULL)
    {
        CComPtr<IStream> spStream;
        hr = ::CreateStreamOnHGlobal(hMem, FALSE, &spStream);
        if (SUCCEEDED(hr))
        {
            hr = m_spDocument->save(CComVariant(spStream));
            if (SUCCEEDED(hr))
            {
                HANDLE hFile = ::CreateFileW(pszFile, GENERIC_WRITE, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
                if (hFile != INVALID_HANDLE_VALUE)
                {
                    LPVOID pXmlData = ::GlobalLock(hMem);
                    if (pXmlData != NULL)
                    {
                        DWORD dwWritten = 0;
                        BOOL fResult = ::WriteFile(hFile, pXmlData, ::GlobalSize(hMem), &dwWritten, NULL);
                        if (fResult)
                        {
                            fResult = ::FlushFileBuffers(hFile);
                            if (!fResult)
                            {
                                hr = HRESULT_FROM_WIN32(::GetLastError());
                            }
                        }
                        else
                        {
                            hr = HRESULT_FROM_WIN32(::GetLastError());
                        }
                        ::GlobalUnlock(hMem);
                    }
                    else
                    {
                        hr = HRESULT_FROM_WIN32(::GetLastError());
                    }
                    ::CloseHandle(hFile);
                }
                else
                {
                    hr = HRESULT_FROM_WIN32(::GetLastError());
                }
            }
        }
        ::GlobalFree(hMem);
    }
    else
    {
        hr = HRESULT_FROM_WIN32(::GetLastError());
    }
    return hr;
}

HRESULT FlushDrive(PCWSTR szOutputFile)
{
    HRESULT hr = S_FALSE;
    int nDrive = PathGetDriveNumber(szOutputFile);
    if (nDrive >= 0)
    {
        WCHAR szPath[MAX_PATH];
        hr = StringCchPrintfW(szPath, ARRAYSIZE(szPath), L"\\\\.\\%lc:", nDrive + L'A');
        if (SUCCEEDED(hr))
        {
            HANDLE hFile = CreateFileW(szPath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
            if (hFile != INVALID_HANDLE_VALUE)
            {
                BOOL fResult = FlushFileBuffers(hFile);
                if (fResult)
                {
                    DWORD dwBytes = 0;
                    fResult = DeviceIoControl(hFile, FSCTL_LOCK_VOLUME, NULL, 0, NULL, 0, &dwBytes, NULL);
                    if (fResult)
                    {
                        fResult = DeviceIoControl(hFile, FSCTL_DISMOUNT_VOLUME, NULL, 0, NULL, 0, &dwBytes, NULL);
                        if (!fResult)
                        {
                            hr = HRESULT_FROM_WIN32(GetLastError());
                        }
                    }
                    else
                    {
                        hr = HRESULT_FROM_WIN32(GetLastError());
                    }
                }
                else
                {
                    hr = HRESULT_FROM_WIN32(GetLastError());
                }
                CloseHandle(hFile);
            }
            else
            {
                hr = HRESULT_FROM_WIN32(GetLastError());
            }
        }
    }
    return hr;
}


void TcxRouteNameDialog::OnSave(WPARAM, LPARAM)
{
    AutoWaitCursor wc;

    HRESULT hr = S_OK;

    // Get output file name
    WCHAR szOutputFile[MAX_PATH] = {};
    if (GetDlgItemTextW(m_hWnd, IDC_EDIT_OUTPUT_FILE, szOutputFile, ARRAYSIZE(szOutputFile)))
    {
        // Get route name
        WCHAR szRouteName[MAX_PATH] = {};
        if (GetDlgItemTextW(m_hWnd, IDC_EDIT_ROUTE_NAME, szRouteName, ARRAYSIZE(szRouteName)))
        {
            // Set route name
            hr = XmlHelpers::SetString(m_spDocument, s_pszXPathNameQuery, szRouteName);
            if (SUCCEEDED(hr))
            {
                // Save the document and flush it to disk
                hr = SaveXmlDocumentAndFlush(szOutputFile);
                if (SUCCEEDED(hr))
                {
                    hr = FlushDrive(szOutputFile);
                    if (SUCCEEDED(hr))
                    {
                        FormatStatusMessage(IDS_STATUS_FILE_SAVED, szOutputFile);
                        FormattedMessageBox(MB_OK | MB_ICONASTERISK, IDS_STATUS_FILE_SAVED, szOutputFile);
                    }
                    else
                    {
                        FormattedMessageBox(MB_ICONERROR | MB_OK, IDS_ERROR_UNABLE_TO_FLUSH_DEVICE, hr);
                    }
                }
                else
                {
                    FormattedMessageBox(MB_ICONERROR | MB_OK, IDS_SAVE_ERROR_SAVING, hr);
                }
            }
            else
            {
                FormattedMessageBox(MB_ICONERROR | MB_OK, IDS_SAVE_ERROR_SETTING_ROUTE_NAME, hr);
            }
        }
        else
        {
            FormattedMessageBox(MB_ICONERROR | MB_OK, IDS_SAVE_ERROR_GETTING_ROUTE_NAME);
        }
    }
    else
    {
        FormattedMessageBox(MB_ICONERROR | MB_OK, IDS_SAVE_ERROR_GETTING_OUTPUT_FILENAME);
    }
}

INT_PTR CALLBACK TcxRouteNameDialog::DialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    BEGIN_MESSAGE_HANDLERS(TcxRouteNameDialog)
    {
        WINDOWS_MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog);
        WINDOWS_MESSAGE_HANDLER(WM_DEVICECHANGE, OnDeviceChange);
        BEGIN_COMMAND_HANDLERS()
        {
            COMMAND_HANDLER_CODE(IDC_EDIT_INPUT_FILE, EN_CHANGE, OnEditChange);
            COMMAND_HANDLER_CODE(IDC_EDIT_OUTPUT_FILE, EN_CHANGE, OnEditChange);
            COMMAND_HANDLER_CODE(IDC_EDIT_ROUTE_NAME, EN_CHANGE, OnEditChange);
            COMMAND_HANDLER(IDCANCEL, OnCancel);
            COMMAND_HANDLER(IDC_SAVE, OnSave);
            COMMAND_HANDLER(IDC_BROWSE_INPUT, OnBrowseInput);
            COMMAND_HANDLER(IDC_BROWSE_OUTPUT, OnBrowseOutput);
            COMMAND_HANDLER(IDC_DELETE_INPUT, OnDeleteInput);
        }
        END_COMMAND_HANDLERS();
    }
    END_MESSAGE_HANDLERS();
}

const WCHAR* TcxRouteNameDialog::s_pszXPathNameQuery = L"/TrainingCenterDatabase/Courses/Course/Name";
const WCHAR* TcxRouteNameDialog::s_pszXPathIdQuery = L"/TrainingCenterDatabase/Folders/Courses/CourseFolder/CourseNameRef/Id";
