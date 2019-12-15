#pragma once

class AutoWaitCursor
{
private:
    HCURSOR m_hOldCursor;

private:
    AutoWaitCursor(const AutoWaitCursor&);
    AutoWaitCursor& operator=(const AutoWaitCursor&);

public:
    AutoWaitCursor(void)
        : m_hOldCursor(SetCursor(LoadCursor(NULL, IDC_WAIT)))
    {
    }
    ~AutoWaitCursor(void)
    {
        SetCursor(m_hOldCursor);
    }
};
