// Copyright 2020 Benbuck Nason
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

#pragma once

// App
#include "ErrorContext.h"

// Windows
#include <Windows.h>
#include <ShObjIdl.h>
#include <comdef.h>

// Standard library
#include <string>
#include <type_traits>

template <class T, class U>
constexpr T narrow_cast(U && u) noexcept
{
    return static_cast<T>(std::forward<U>(u));
}

HINSTANCE getInstance() noexcept;
std::string getResourceString(unsigned int id);
bool isWindowStealth(HWND hwnd) noexcept;
bool isWindowUserVisible(HWND hwnd) noexcept;
void errorMessage(unsigned int id);
void errorMessage(const ErrorContext & errorContext);

// RAII wrapper that creates the IVirtualDesktopManager once and
// caches it for the lifetime of the object (typically your app).
class VirtualDesktopHelper
{
public:
    VirtualDesktopHelper() : m_comReady(false), m_pVDM(nullptr) {}

    ~VirtualDesktopHelper()
    {
        if (m_pVDM)
            m_pVDM->Release();

        if (m_comReady)
            ::CoUninitialize();
    }

    VirtualDesktopHelper(const VirtualDesktopHelper&) = delete;
    VirtualDesktopHelper& operator=(const VirtualDesktopHelper&) = delete;

    bool Init()
    {
        HRESULT hr = ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        m_comReady = SUCCEEDED(hr);
        if (FAILED(hr) && hr == RPC_E_CHANGED_MODE)
            m_comReady = true;

        if (!m_comReady)
            return false;

        hr = ::CoCreateInstance(
            CLSID_VirtualDesktopManager,
            nullptr,
            CLSCTX_ALL,
            IID_IVirtualDesktopManager,
            reinterpret_cast<void**>(&m_pVDM));

        return SUCCEEDED(hr) && m_pVDM;
    }

    bool IsOnCurrentDesktop(HWND hwnd) const
    {
        if (!m_pVDM || !hwnd || !::IsWindow(hwnd))
            return false;

        if (!::IsWindowVisible(hwnd))
            return false;

        // Get the window's desktop GUID via COM.
        // (This is a property of the window itself, so it doesn't
        // suffer from the same staleness problem.)
        GUID windowDesktopId = {};
        HRESULT hr = m_pVDM->GetWindowDesktopId(hwnd, &windowDesktopId);
        if (FAILED(hr))
            return false;

        // A zero GUID means the window is pinned to all desktops
        // (e.g. sticky notes, Task Manager when "show on all desktops").
        if (::IsEqualGUID(windowDesktopId, GUID_NULL))
            return true;

        // Read the current desktop GUID straight from the registry.
        // This value is updated by Explorer instantly on switch,
        // before the COM proxy catches up.
        GUID currentDesktopId = {};
        if (!GetCurrentDesktopIdFromRegistry(currentDesktopId))
            return false;

        return ::IsEqualGUID(windowDesktopId, currentDesktopId);
    }

private:
    static bool GetCurrentDesktopIdFromRegistry(GUID& desktopId)
    {
        HKEY hKey = nullptr;
        LONG rc = ::RegOpenKeyExW(
            HKEY_CURRENT_USER,
            L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops",
            0,
            KEY_READ,
            &hKey);

        if (rc != ERROR_SUCCESS)
            return false;

        BYTE data[16] = {};   // A GUID is exactly 16 bytes
        DWORD dataSize = sizeof(data);
        DWORD type = 0;

        rc = ::RegQueryValueExW(
            hKey,
            L"CurrentVirtualDesktop",
            nullptr,
            &type,
            data,
            &dataSize);

        ::RegCloseKey(hKey);

        if (rc != ERROR_SUCCESS || type != REG_BINARY || dataSize != 16)
            return false;

        ::memcpy(&desktopId, data, 16);
        return true;
    }

    bool                    m_comReady;
    IVirtualDesktopManager* m_pVDM;
};

extern VirtualDesktopHelper g_vdHelper;
