// header.h : include file for standard system include files,
// or project specific include files
//

#pragma once

#include "targetver.h"
#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files
#include <windows.h>
#include <windowsx.h>
#include <objbase.h>
#include <atlbase.h>
#include <strsafe.h>
#include <commdlg.h>
#include <commctrl.h>
#include <stdlib.h>
#include <tchar.h>
#include <dbt.h>

#define HINST_THIS_MODULE reinterpret_cast<HINSTANCE>(&__ImageBase)