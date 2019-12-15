#pragma once

#include <windows.h>
#include <atlbase.h>

class TcxRouteNameDialog
{
private:
    HWND                     m_hWnd;
    CComPtr<IXMLDOMDocument> m_spDocument;
    bool                     m_fIgnoreDeviceChange;
    static const WCHAR* s_pszXPathIdQuery;
    static const WCHAR* s_pszXPathNameQuery;

private:
    explicit TcxRouteNameDialog(HWND hWnd);
    ~TcxRouteNameDialog();

    void ReplaceCharactersInPlace(PWSTR pszString, WCHAR chFrom, WCHAR chTo);
    BOOL SetupDeviceArrivalNotifications(IN GUID InterfaceClassGuid, OUT HDEVNOTIFY* hDeviceNotify);
    void UpdateSaveButtonState();
    int DisplayFormattedMessage(int nButtons, UINT uFormatStringId, ...);
    HRESULT GetFirstGarminDeviceNewFilesDirectory(PWSTR pszDrive, size_t cchDrive);
    HRESULT LoadFile(PCWSTR pszFile);
    void ConstructOutputFileName();

    LRESULT OnInitDialog(WPARAM, LPARAM);
    LRESULT OnDeviceChange(WPARAM wParam, LPARAM lParam);

    void OnCancel(WPARAM, LPARAM);
    void OnEditChange(WPARAM, LPARAM);
    void OnBrowseInput(WPARAM, LPARAM);
    void OnBrowseOutput(WPARAM, LPARAM);
    void OnSave(WPARAM, LPARAM);

public:
    static INT_PTR CALLBACK DialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
};
