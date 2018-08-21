// mscoree.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"

#include <cassert>
#include <mutex>
#include <vector>
#include <string>

#include <atlbase.h>
#include <metahost.h>
#include <mscoree.h>

using namespace ATL;

#define WCHAR wchar_t
#define W(str) L ## str
#define RETURN_IF_FAILED(exp) { hr = (exp); if (FAILED(hr)) { assert(false && #exp); return hr; } }
#define RETURN_VOID_IF_FAILED(exp) { hr = (exp); if (FAILED(hr)) { assert(false && #exp); return; } }

//
// Declaration of MSCOREE forwarders - undocumented
//
STDAPI CreateConfigStream(void *, void *); // CLR startup
STDAPI Ordinal24(void *, void *); // CLR startup
STDAPI_(void *) GetProcessExecutableHeap(); // Debugging

namespace
{
    class InitCom final
    {
    public:
        InitCom()
            : Result{ ::CoInitializeEx(nullptr, COINITBASE_MULTITHREADED) }
        {
            assert(SUCCEEDED(Result));
        }

        ~InitCom()
        {
            if (SUCCEEDED(Result))
                ::CoUninitialize();
        }

        const HRESULT Result;
    };

    HRESULT GetActivationAssemblies(_Inout_ std::vector<std::wstring> &assemblies, _Inout_ std::vector<const WCHAR *> &assembliesRaw)
    {
        assert(assemblies.empty());

        SIZE_T dataSize;

        // This data block is re-used multiple times.
        // Using the maximum possible size that could ever be needed in this function.
        BYTE data[sizeof(ACTIVATION_CONTEXT_ASSEMBLY_DETAILED_INFORMATION) + (MAX_PATH * 4)];

        // Get current context details
        BOOL suc = ::QueryActCtxW(
            QUERY_ACTCTX_FLAG_USE_ACTIVE_ACTCTX,
            nullptr,
            nullptr,
            ActivationContextDetailedInformation,
            data,
            ARRAYSIZE(data),
            &dataSize);
        if (suc == FALSE)
            return HRESULT_FROM_WIN32(::GetLastError());

        auto cdi = reinterpret_cast<PACTIVATION_CONTEXT_DETAILED_INFORMATION>(data);
        const DWORD assemblyCount = cdi->ulAssemblyCount;

        try
        {
            // Assembly index '0' is reserved and count doesn't include the reserved entry.
            for (DWORD adidx = 1; adidx <= assemblyCount; ++adidx)
            {
                // Collect all assemblies involved in the current context
                suc = ::QueryActCtxW(
                    QUERY_ACTCTX_FLAG_USE_ACTIVE_ACTCTX,
                    nullptr,
                    &adidx,
                    AssemblyDetailedInformationInActivationContext,
                    data,
                    ARRAYSIZE(data),
                    &dataSize);
                if (suc == FALSE)
                    return HRESULT_FROM_WIN32(::GetLastError());

                auto adi = reinterpret_cast<PACTIVATION_CONTEXT_ASSEMBLY_DETAILED_INFORMATION>(data);
                if (adi->ulManifestPathLength > 0)
                {
                    assemblies.push_back({ adi->lpAssemblyManifestPath });
                    assembliesRaw.push_back(assemblies.back().c_str());
                }
            }
        }
        catch (const std::bad_alloc&)
        {
            return E_OUTOFMEMORY;
        }

        return S_OK;
    }

    class ClrInstance final
    {
    public: // static
        static decltype(&::CreateConfigStream) CreateConfigStream;
        static decltype(&::Ordinal24) Ordinal24;
        static decltype(&::GetProcessExecutableHeap) GetProcessExecutableHeap;

    private: //static
        static std::once_flag LoadClrOnce;
        static HRESULT LoadResult;
        static CComPtr<ICLRRuntimeHost> RuntimeHost;
        static void LoadClr(_In_z_ const WCHAR *version)
        {
            assert(version != nullptr);

            // The reference will be set in the 'RETURN' macros
            HRESULT &hr = LoadResult;

            WCHAR pathBuffer[MAX_PATH];
            (void)::ExpandEnvironmentStringsW(W("%windir%\\system32\\mscoree.dll"), pathBuffer, ARRAYSIZE(pathBuffer));

            // This call is relying on file system redirection for the computed path above.
            // https://docs.microsoft.com/en-us/windows/desktop/winprog64/file-system-redirector
            HMODULE hmod = ::LoadLibraryExW(pathBuffer, nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
            assert(hmod != NULL);

            // Get functions that require forwarding to the real MSCOREE.
            CreateConfigStream = (decltype(CreateConfigStream))::GetProcAddress(hmod, "CreateConfigStream");
            assert(CreateConfigStream != nullptr);
            Ordinal24 = (decltype(Ordinal24))::GetProcAddress(hmod, MAKEINTRESOURCEA(24));
            assert(Ordinal24 != nullptr);
            GetProcessExecutableHeap = (decltype(GetProcessExecutableHeap))::GetProcAddress(hmod, "GetProcessExecutableHeap");
            assert(GetProcessExecutableHeap != nullptr);

            // Get the create CLR instance entry point
            auto createInstance = (decltype(&::CLRCreateInstance))::GetProcAddress(hmod, "CLRCreateInstance");
            assert(createInstance != nullptr);

            CComPtr<ICLRMetaHost> metaHost;
            RETURN_VOID_IF_FAILED(createInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, (void**)&metaHost));

            CComPtr<ICLRRuntimeInfo> runtimeInfo;
            RETURN_VOID_IF_FAILED(metaHost->GetRuntime(version, IID_ICLRRuntimeInfo, (void**)&runtimeInfo));

            CComPtr<ICLRRuntimeHost> runtimeHost;
            RETURN_VOID_IF_FAILED(runtimeInfo->GetInterface(CLSID_CLRRuntimeHost, IID_ICLRRuntimeHost, (void**)&runtimeHost));

            // Start the runtime
            RETURN_VOID_IF_FAILED(runtimeHost->Start());

            // Transfer ownership
            RuntimeHost.Attach(runtimeHost.Detach());

            // Ensure the result is success
            LoadResult = S_OK;
        }

    public:
        ClrInstance()
        {
            const WCHAR *version = W("v4.0.30319");
            std::call_once(LoadClrOnce, LoadClr, version);
            assert(SUCCEEDED(LoadResult));
        }

        HRESULT GetClassFactoryForType(_In_ REFCLSID rclsid, _In_ REFIID riid, _Out_ LPVOID* ppv)
        {
            HRESULT hr;
            RETURN_IF_FAILED(LoadResult);
            assert(RuntimeHost != nullptr);

            // Get all activation assemblies
            std::vector<std::wstring> assemblies;
            std::vector<const WCHAR *> assembliesRaw;
            RETURN_IF_FAILED(GetActivationAssemblies(assemblies, assembliesRaw));

            // Wrap the returned CCW in a smart pointer since
            // `Marshal.GetIUnknownForObject()` is used in managed code
            // and it performs an `AddRef()`.
            CComPtr<IUnknown> ccw;

            // See LoaderShim for matching struct layout in managed
            struct ActivationRequest
            {
                GUID ClassId;
                GUID InterfaceId;
                DWORD AssemblyCount;
                const WCHAR **AssemblyList;
                void **ClassFactoryDest;
            } req{ rclsid, riid, static_cast<DWORD>(assembliesRaw.size()), assembliesRaw.data(), (void**)&ccw };

            // Convert the address of the request to a string
            WCHAR ptrAsString[ARRAYSIZE(W("9223372036854775807"))];
            ::swprintf_s(ptrAsString, W("%zu"), (uintptr_t)&req);

            HRESULT activationResult;
            RETURN_IF_FAILED(RuntimeHost->ExecuteInDefaultAppDomain(
                W("LoaderShim.dll"),    // In current working directory
                W("LoaderShim.ComActivator"),
                W("GetClassFactoryForType"),
                ptrAsString,
                (DWORD *)&activationResult));

            if (FAILED(activationResult))
                return activationResult;

            // Query CCW for requested interface
            return ccw->QueryInterface(riid, ppv);
        }
    };

    decltype(&::CreateConfigStream) ClrInstance::CreateConfigStream;
    decltype(&::Ordinal24) ClrInstance::Ordinal24;
    decltype(&::GetProcessExecutableHeap) ClrInstance::GetProcessExecutableHeap;

    std::once_flag ClrInstance::LoadClrOnce;
    HRESULT ClrInstance::LoadResult;
    CComPtr<ICLRRuntimeHost> ClrInstance::RuntimeHost;
}

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Out_ LPVOID FAR* ppv)
{
    HRESULT hr;

    InitCom initCom;
    RETURN_IF_FAILED(initCom.Result);

    ClrInstance clrInstance;
    RETURN_IF_FAILED(clrInstance.GetClassFactoryForType(rclsid, riid, ppv));

    return S_OK;
}

STDAPI DllCanUnloadNow()
{
    return S_FALSE;
}

//
// Definition of MSCOREE forwarders
//

STDAPI CreateConfigStream(void *a, void *b)
{
    return ClrInstance::CreateConfigStream(a, b);
}

STDAPI Ordinal24(void *a, void *b)
{
    return ClrInstance::Ordinal24(a, b);
}

STDAPI_(void *) GetProcessExecutableHeap()
{
    return ClrInstance::GetProcessExecutableHeap();
}
