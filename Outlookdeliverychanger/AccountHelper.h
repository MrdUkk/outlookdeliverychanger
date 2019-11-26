/*
	This is derivative work based on article:
	"Using Account Management API (IOlkAccountManger) to List Outlook Email Accounts"
	by Ashutosh Bhawasinka,  28 Aug 2008
	published first on www.codeproject.com

	(c) Wholesome Software 2011-2019
*/
#pragma once

#include <initguid.h>
#include <mapiguid.h>

#include "AcctMgmt.h"

class CAccountHelper : public IOlkAccountHelper
{
public:
	CAccountHelper(LPWSTR lpwszProfName, LPMAPISESSION lpSession);
	~CAccountHelper();

	// IUnknown
	STDMETHODIMP			QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();

	// IOlkAccountHelper
	inline STDMETHODIMP PlaceHolder1(LPVOID)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP GetIdentity(LPWSTR pwszIdentity, DWORD * pcch);
	STDMETHODIMP GetMapiSession(LPUNKNOWN * ppmsess);
	STDMETHODIMP HandsOffSession();

private:
	LONG		m_cRef;
	LPUNKNOWN	m_lpUnkSession;
	LPWSTR		m_lpwszProfile;
	size_t		m_cchProfile;
};