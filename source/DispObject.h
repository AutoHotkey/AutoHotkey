/*
DispObject.h

Original code by Steve Gray.

This software is provided 'as-is', without any express or implied
warranty. In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, without restriction.
*/

#pragma once


// DispObject : Template for boiler-plate IDispatch objects.
template <typename I>
class DispObject : public I {
public:
	// IUnknown
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();
	STDMETHOD(QueryInterface)(REFIID iid, void** ppv);

	// IDispatch
	STDMETHOD(GetTypeInfoCount)(UINT* pCountTypeInfo);
	STDMETHOD(GetTypeInfo)(UINT iTypeInfo, LCID lcid, ITypeInfo** ppITypeInfo);
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
	STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams
					, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);
	
	DispObject(ITypeInfo *typeInfo) : m_cRef(1), m_pTypeInfo(typeInfo) {}
	~DispObject();

	typedef I *DispInterface;

private:
	ULONG m_cRef;
	ITypeInfo* m_pTypeInfo;
};


HRESULT LoadMyTypeInfo(REFIID riid, ITypeInfo **ppTypeInfo);


template<class T, typename I>
HRESULT CreateDispatchInstance(I **ppInst)
{
	ITypeInfo *typeInfo;
	auto hr = LoadMyTypeInfo(__uuidof(T::DispInterface), &typeInfo);
	if (FAILED(hr))
	{
		*ppInst = nullptr;
		return hr;
	}
	*ppInst = new T(typeInfo);
	return S_OK;
}


template<typename I>
DispObject<I>::~DispObject()
{
	if (m_pTypeInfo)
		m_pTypeInfo->Release();
}


//
// IUnknown methods
//

template<typename I>
STDMETHODIMP DispObject<I>::QueryInterface(REFIID iid, void ** ppv)
{
	if (iid == IID_IUnknown)
		*ppv = static_cast<IUnknown*>(this);
	else if (iid == __uuidof(I))
		*ppv = static_cast<I*>(this);
	else if (iid == IID_IDispatch)
		*ppv = static_cast<IDispatch*>(this);
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

template<typename I>
STDMETHODIMP_(ULONG) DispObject<I>::AddRef()
{
	return ++m_cRef;
}

template<typename I>
STDMETHODIMP_(ULONG) DispObject<I>::Release()
{
	if (--m_cRef != 0)
		return m_cRef;
	delete this;
	return 0;
}


//
// IDispatch methods
//

template<typename I>
STDMETHODIMP DispObject<I>::GetTypeInfoCount(UINT * pCountTypeInfo)
{
	*pCountTypeInfo = 1;
	return S_OK;
}

template<typename I>
STDMETHODIMP DispObject<I>::GetTypeInfo(UINT iTypeInfo, LCID lcid, ITypeInfo ** ppITypeInfo)
{
	if (iTypeInfo != 0)
		return DISP_E_BADINDEX;
	m_pTypeInfo->AddRef();
	*ppITypeInfo = m_pTypeInfo;
	return S_OK;
}

template<typename I>
STDMETHODIMP DispObject<I>::GetIDsOfNames(REFIID riid, LPOLESTR * rgszNames, UINT cNames, LCID lcid, DISPID * rgDispId)
{
	return m_pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

template<typename I>
STDMETHODIMP DispObject<I>::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult, EXCEPINFO * pExcepInfo, UINT * puArgErr)
{
	return m_pTypeInfo->Invoke(static_cast<I*>(this), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}