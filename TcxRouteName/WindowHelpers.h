#pragma once

#include <windows.h>

namespace WindowHelpers
{
    inline int RectWidth(const RECT& rc)
    {
        return (rc.right - rc.left);
    }

    inline int RectHeight(const RECT& rc)
    {
        return (rc.bottom - rc.top);
    }

    void CenterWindow(HWND hWnd, HWND hWndParent)
    {
        // If our parent is NULL minimized, use the desktop window
        if (!hWndParent || (GetWindowLong(hWndParent, GWL_STYLE) & WS_MINIMIZE))
        {
            hWndParent = GetDesktopWindow();
        }

        // Get the window rects
        RECT rcParent;
        GetWindowRect(hWndParent, &rcParent);

        RECT rcCurrent;
        GetWindowRect(hWnd, &rcCurrent);

        // Get the desired coordinates for the upper-left hand corner
        RECT rcFinal;
        rcFinal.left = rcParent.left + (RectWidth(rcParent) - RectWidth(rcCurrent)) / 2;
        rcFinal.top = rcParent.top + (RectHeight(rcParent) - RectHeight(rcCurrent)) / 2;
        rcFinal.right = rcFinal.left + RectWidth(rcCurrent);
        rcFinal.bottom = rcFinal.top + RectHeight(rcCurrent);

        // Make sure we're not off the screen
        HMONITOR hMonitor = MonitorFromRect(&rcFinal, MONITOR_DEFAULTTONEAREST);
        if (hMonitor)
        {
            // Get the screen coordinates of this monitor
            MONITORINFO MonitorInfo = { 0 };
            MonitorInfo.cbSize = sizeof(MonitorInfo);
            if (GetMonitorInfo(hMonitor, &MonitorInfo))
            {
                // Ensure the window is in the working area's region
                rcFinal.left = max(MonitorInfo.rcWork.left, min(MonitorInfo.rcWork.right - RectWidth(rcCurrent), rcFinal.left));
                rcFinal.top = max(MonitorInfo.rcWork.top, min(MonitorInfo.rcWork.bottom - RectHeight(rcCurrent), rcFinal.top));
            }
        }

        // Move it
        SetWindowPos(hWnd, NULL, rcFinal.left, rcFinal.top, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER);
    }
}
