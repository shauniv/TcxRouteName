// TcxRouteName.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "resource.h"
#include "TcxRouteNameDialog.h"

int CALLBACK wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE, _In_ LPWSTR, _In_ int)
{
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr))
    {
        InitCommonControls();
        DialogBox(hInstance, MAKEINTRESOURCE(IDD_TCX_INPUT_DIALOG), NULL, TcxRouteNameDialog::DialogProc);
        CoUninitialize();
    }
    return 0;
}
