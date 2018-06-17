#pragma once

#include <Windows.h>
#include <atlbase.h>

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
        if (pUnkOuter != nullptr)
            return CLASS_E_NOAGGREGATION;

        try
        {
            ATL::CComPtr<T> inst;
            inst.Attach(new T());
            return inst->QueryInterface(riid, ppvObject);
        }
        catch (const std::bad_alloc &)
        {
            return E_OUTOFMEMORY;
        }
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
