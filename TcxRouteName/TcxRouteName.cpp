// TcxRouteName.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "resource.h"
#include "TcxRouteNameDialog.h"

int CALLBACK wWinMain(_In_ HINSTANCE, _In_opt_ HINSTANCE, _In_ LPWSTR, _In_ int)
{
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr))
    {
        INITCOMMONCONTROLSEX icex = {};
        icex.dwSize = sizeof(icex);
        icex.dwICC = ICC_BAR_CLASSES;
        InitCommonControlsEx(&icex);
        DialogBox(HINST_THIS_MODULE, MAKEINTRESOURCE(IDD_TCX_INPUT_DIALOG), NULL, TcxRouteNameDialog::DialogProc);
        CoUninitialize();
    }
    return 0;
}
