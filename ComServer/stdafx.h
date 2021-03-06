#pragma once

#define WIN32_LEAN_AND_MEAN

// Windows Header Files:
#include <windows.h>

#define _ATL_NO_AUTOMATIC_NAMESPACE
#include <atlbase.h>
#include <atlstr.h>

#define RETURN_IF_FAILED(exp) { hr = exp; if (FAILED(hr)) { return hr; } }
