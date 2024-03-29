#pragma once

#include <windows.h>
#include <windowsx.h>
#include <shlobj.h>
#include <atlbase.h>

class TcxRouteNameDialog
{
private:
    HWND                     m_hWnd;
    CComPtr<IXMLDOMDocument> m_spDocument;
    bool                     m_fIgnoreDeviceChange;
    bool                     m_fInputFileDeleted;
    static const WCHAR* s_pszXPathIdQuery;
    static const WCHAR* s_pszXPathNameQuery;

private:
    explicit TcxRouteNameDialog(HWND hWnd);
    ~TcxRouteNameDialog();

    void ReplaceCharactersInPlace(PWSTR pszString, WCHAR chFrom, WCHAR chTo);
    BOOL SetupDeviceArrivalNotifications(IN GUID InterfaceClassGuid, OUT HDEVNOTIFY* hDeviceNotify);
    void UpdateSaveButtonState();
    void UpdateDeleteButtonState();
    int FormattedMessageBox(int nButtons, UINT uFormatStringId, ...);
    void FormatStatusMessage(UINT uFormatStringId, ...);
    HRESULT GetAllGarminDeviceNewFilesDirectories(PWSTR* ppszPaths);
    HRESULT LoadXmlDocument(PCWSTR pszFile, IXMLDOMDocument** ppXmlDocument);
    HRESULT ValidateTcxRoute(IXMLDOMDocument* pXmlDocument);
    HRESULT LoadFile(PCWSTR pszFile);
    void AddGarminDevicesToDirectoryList();
    void ConstructOutputPathName();
    HRESULT SaveXmlDocumentAndFlush(PCWSTR pszFile);
    HRESULT GetFileCount(PCWSTR pszDirectory, PCWSTR pszPattern, int* nFileCount);
    HRESULT CountRoutesOnDevice(PCWSTR pszOutputFilename, int* pnNewFilesCount, int* pnCoursesCount);

    LRESULT OnInitDialog(WPARAM, LPARAM);
    LRESULT OnDeviceChange(WPARAM wParam, LPARAM lParam);

    void OnCancel(WPARAM, LPARAM);
    void OnEditChange(WPARAM, LPARAM);
    void OnBrowseInput(WPARAM, LPARAM);
    void OnBrowseOutput(WPARAM, LPARAM);
    void OnDirectorySelChange(WPARAM, LPARAM);
    void OnDeleteInput(WPARAM, LPARAM);
    void OnSave(WPARAM, LPARAM);

    static int CALLBACK BrowseCallbackProc(HWND hWnd, UINT uMsg, LPARAM, LPARAM pData);

public:
    static INT_PTR CALLBACK DialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
};
