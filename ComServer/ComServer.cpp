#include "stdafx.h"
#include "ComServer.h"

HRESULT STDMETHODCALLTYPE Outer::OuterImpl::ComputeFibonacci(_In_ int n, _Out_ int *fib)
{
    if (fib == nullptr)
        return E_POINTER;

    if (n < 0 || 46 < n)
        return E_INVALIDARG;

    if (n == 0 || n == 1)
    {
        *fib = n;
    }
    else
    {
        int p = 0;
        int c = 1;

        while (n >= 2)
        {
            int t = c;
            c = c + p;
            p = t;

            --n;
        }

        *fib = c;
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE Server::ComputePi(_Out_ double *pi)
{
    if (pi == nullptr)
        return E_POINTER;

    double sum = 0.0;
    int sign = 1;
    for (int i = 0; i < 1024; ++i)
    {
        sum += sign / (2.0 * i + 1.0);
        sign *= -1;
    }

    *pi = 4.0 * sum;
    return S_OK;
}
