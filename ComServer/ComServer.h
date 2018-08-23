#pragma once

#include <Windows.h>
#include <atlbase.h>
#include <exception>

template<typename T>
inline HRESULT CreateInstanceImpl(_In_opt_ IUnknown *outer, _In_  REFIID riid, _Outptr_ T **inst)
{
    if (inst == nullptr)
        return E_POINTER;

    if (outer != nullptr)
        return CLASS_E_NOAGGREGATION;

    try
    {
        *inst = new T();
    }
    catch (const std::bad_alloc&)
    {
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

// IOuter definition
DECLARE_INTERFACE_IID_(IOuter, IUnknown, "575375A2-3F92-44A8-89A3-A7BB87BE9622")
{
    // Compute and return the request Fibonacci number
    STDMETHOD(ComputeFibonacci)(_In_ int n, _Out_ int *fib) PURE;
};

// Server implementation of IOuter
class DECLSPEC_UUID("BCF98F86-300D-4245-82F2-3D9D1E62B1FC") Outer : IUnknown
{
public: // Outer
    Outer(_In_opt_ IUnknown *outer)
        : _pUnkOuter{ outer }
        , _iOuterImpl{}
    {
        _iOuterImpl._pUnkOuter = outer;
    }

private: // impl
    class OuterImpl : IOuter
    {
        friend class Outer;
        IUnknown *_pUnkOuter;
    public: // IOuter
        STDMETHOD(ComputeFibonacci)(_In_ int n, _Out_ int *fib);

    public: // IUnknown
        STDMETHODIMP QueryInterface(REFIID riid, void** ppv) { return _pUnkOuter->QueryInterface(riid, ppv); };
        STDMETHOD_(ULONG, AddRef)(void) { return _pUnkOuter->AddRef(); };
        STDMETHOD_(ULONG, Release)(void) { return _pUnkOuter->Release(); };
    };

    OuterImpl _iOuterImpl;

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
        else if (riid == __uuidof(IOuter))
        {
            *ppvObject = &_iOuterImpl;
        }
        else
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        ((IUnknown*)*ppvObject)->AddRef();
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
    IUnknown* _pUnkOuter;
    ULONG _refCount = 1;
};

template<>
inline HRESULT CreateInstanceImpl<Outer>(_In_opt_ IUnknown *outer, _In_  REFIID riid, _Outptr_ Outer **inst)
{
    if (inst == nullptr)
        return E_POINTER;

    if (outer != nullptr && riid != IID_IUnknown)
        return CLASS_E_NOAGGREGATION;

    try
    {
        *inst = new Outer(outer);
    }
    catch (const std::bad_alloc&)
    {
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

// IServer definition
DECLARE_INTERFACE_IID_(IServer, IUnknown, "F38720E5-2D64-445E-88FB-1D696F614C78")
{
    // Compute and return the value of PI
    STDMETHOD(ComputePi)(_Out_ double *pi) PURE;
};

// Server implementation of IServer
class DECLSPEC_UUID("114383E9-1969-47D2-9AA9-91388C961A19") Server : IServer
{
public: // IServer
    STDMETHOD(ComputePi)(_Out_ double *pi);

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
        else if (riid == __uuidof(IServer))
        {
            *ppvObject = static_cast<IServer *>(this);
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

// Templated class factory
template<typename T>
class ClassFactory : IClassFactory
{
public: // static
    static HRESULT Create(_In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
    {
        try
        {
            ATL::CComPtr<ClassFactory> inst;
            inst.Attach(new ClassFactory());
            return inst->QueryInterface(riid, ppv);
        }
        catch (const std::bad_alloc &)
        {
            return E_OUTOFMEMORY;
        }
    }

public: // IClassFactory
    STDMETHOD(CreateInstance)(
        _In_opt_  IUnknown *pUnkOuter,
        _In_  REFIID riid,
        _COM_Outptr_  void **ppvObject)
    {
        ATL::CComPtr<T> inst;
        HRESULT hr = CreateInstanceImpl<T>(pUnkOuter, riid, &inst);
        if (FAILED(hr))
            return hr;

        return inst->QueryInterface(riid, ppvObject);
    }

    STDMETHOD(LockServer)(/* [in] */ BOOL fLock)
    {
        return ::CoLockObjectExternal(static_cast<IUnknown *>(this), fLock, TRUE);
    }

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
        else if (riid == __uuidof(IClassFactory))
        {
            *ppvObject = static_cast<IClassFactory *>(this);
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
