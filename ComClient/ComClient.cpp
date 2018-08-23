#include "stdafx.h"
#include <Windows.h>
#include <atlbase.h>

using namespace ATL;

// IOuter definition
DECLARE_INTERFACE_IID_(IOuter, IUnknown, "575375A2-3F92-44A8-89A3-A7BB87BE9622")
{
    // Compute and return the request Fibonacci number
    STDMETHOD(ComputeFibonacci)(_In_ int n, _Out_ int *fib) PURE;
};

// IServer definition
DECLARE_INTERFACE_IID_(IServer, IUnknown, "F38720E5-2D64-445E-88FB-1D696F614C78")
{
    // Compute and return the value of PI
    STDMETHOD(ComputePi)(_Out_ double *pi) PURE;
};

const IID IID_IServer = __uuidof(IServer);

// {114383E9-1969-47D2-9AA9-91388C961A19}
const CLSID CLSID_Server = { 0x114383E9, 0x1969, 0x47D2, { 0x9A, 0xA9, 0x91, 0x38, 0x8C, 0x96, 0x1A, 0x19 } };

HRESULT CallOuter(_In_ IUnknown *o, _Outptr_ IOuter **r)
{
    CComPtr<IOuter> outer;
    HRESULT hr = o->QueryInterface(&outer);
    if (FAILED(hr))
        return hr;

    int fibn = 10;
    int fibResult;
    hr = outer->ComputeFibonacci(fibn, &fibResult);
    if (FAILED(hr))
        return hr;

    ::printf("fib(%d) = %d\n", fibn, fibResult);
    *r = outer.Detach();
    return S_OK;
}

HRESULT CallServer(_In_ IUnknown *s, _Outptr_ IServer **r)
{
    CComPtr<IServer> server;
    HRESULT hr = s->QueryInterface(&server);
    if (FAILED(hr))
        return hr;

    double pi;
    hr = server->ComputePi(&pi);
    if (FAILED(hr))
        return hr;

    ::printf("\u03C0 = %f\n", pi);
    *r = server.Detach();
    return S_OK;
}

HRESULT QueryServer()
{
    CComPtr<IUnknown> inst;
    HRESULT hr = ::CoCreateInstance(CLSID_Server, nullptr, CLSCTX_INPROC, IID_IUnknown, (void**)&inst);
    if (FAILED(hr))
        return hr;

    CComPtr<IServer> server;
    hr = CallServer(inst, &server);
    if (FAILED(hr))
        return hr;

    CComPtr<IOuter> outer;
    hr = CallOuter(inst, &outer);
    if (FAILED(hr))
        return hr;

    // Check for proper aggregation - IUnknowns should match
    CComPtr<IUnknown> outerUnk;
    hr = outer.QueryInterface(&outerUnk);
    if (FAILED(hr))
        return hr;

    CComPtr<IUnknown> outerUnkMaybe;
    hr = server.QueryInterface(&outerUnkMaybe);
    if (FAILED(hr))
        return hr;

    if (outerUnk.p != outerUnkMaybe.p)
        return E_UNEXPECTED;

    return S_OK;
}

class OuterServer : public IUnknown
{
public: // IUnknown
    STDMETHOD(QueryInterface)(
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
    {
        if (ppvObject == nullptr)
            return E_POINTER;

        if (riid == __uuidof(IUnknown))
        {
            *ppvObject = static_cast<IUnknown *>(this);
        }
        else
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    STDMETHOD_(ULONG, AddRef)(void)
    {
        return ::InterlockedIncrement(&_refCount);
    }

    STDMETHOD_(ULONG, Release)(void)
    {
        ULONG c = ::InterlockedDecrement(&_refCount);
        if (c == 0)
            delete this;
        return c;
    }

private:
    ULONG _refCount = 1;
};

HRESULT QueryServer_Aggregation()
{
    CComPtr<OuterServer> outerServer;
    outerServer.Attach(new OuterServer());

    CComPtr<IUnknown> inst;
    HRESULT hr = ::CoCreateInstance(CLSID_Server, outerServer, CLSCTX_INPROC, IID_IUnknown, (void**)&inst);
    if (FAILED(hr))
        return hr;

    CComPtr<IServer> server;
    hr = CallServer(inst, &server);
    if (FAILED(hr))
        return hr;

    CComPtr<IOuter> outer;
    hr = CallOuter(inst, &outer);
    if (FAILED(hr))
        return hr;

    // Check for proper aggregation - IUnknowns should match
    CComPtr<IUnknown> outerUnk;
    hr = outerServer.QueryInterface(&outerUnk);
    if (FAILED(hr))
        return hr;

    CComPtr<IUnknown> outerUnkMaybe1;
    hr = outer.QueryInterface(&outerUnkMaybe1);
    if (FAILED(hr))
        return hr;

    CComPtr<IUnknown> outerUnkMaybe2;
    hr = server.QueryInterface(&outerUnkMaybe2);
    if (FAILED(hr))
        return hr;

    if (outerUnk.p != outerUnkMaybe1.p
        || outerUnk.p != outerUnkMaybe2.p)
        return E_UNEXPECTED;

    return S_OK;
}

int main()
{
    // Set console codepage to utf-8. Also note the '/utf-8' compiler flag.
    ::SetConsoleOutputCP(65001);

    ::CoInitializeEx(0, COINITBASE_MULTITHREADED);

    HRESULT hr = QueryServer();

    ::CoUninitialize();

    if (FAILED(hr))
    {
        ::printf("Failure: 0x%08x\n", hr);
        return EXIT_FAILURE;
    }

    return EXIT_SUCCESS;
}
