/*
	This is derivative work based on article:
	"Using Account Management API (IOlkAccountManger) to List Outlook Email Accounts"
	by Ashutosh Bhawasinka,  28 Aug 2008
	published first on www.codeproject.com

	(c) Wholesome Software 2011-2019
*/
#include "stdafx.h"
#include <MAPIUTIL.H>
#include "AccountHelper.h"


#define PST_EXTERN_PROPID_BASE (0x6700)
#define PR_PST_PATH PROP_TAG(PT_STRING8, PST_EXTERN_PROPID_BASE + 0)
#define PROP_MAPI_SERVICE_UID PROP_TAG(PT_BINARY, 0x2000)

CMAPIWrapper::CMAPIWrapper()
{
	hResErr=MAPIInitialize(NULL);
	if(FAILED(hResErr)) Error=true;
	else Error=false;
}

CMAPIWrapper::~CMAPIWrapper()
{
	//need to gracefully shutdown our internal data
	//<HERE>
	for(int a=0;a<ProfilesList.GetSize();a++)
	{
		CMAPIProfile *profile=(CMAPIProfile *)ProfilesList.GetAt(a);
		for(int b=0;b<profile->StoresList.GetSize();b++)
		{
			CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(b);
			if(store->cbStoreName > 0) free(store->StoreName);
			if(store->cbServiceUID > 0) free(store->ServiceUID);
			if(store->cbEntryIDInbox > 0) free(store->EntryIDInbox);
			if(store->cbEntryIDstore > 0) free(store->EntryIDstore);
			delete store;
		}
		for(int c=0;c<profile->AccountsList.GetSize();c++)
		{
			CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(c);
			if(account->cbDeliveryFolder > 0) free(account->DeliveryFolder);
			if(account->cbDeliveryStore > 0) free(account->DeliveryStore);
			delete account;
		}
		profile->StoresList.RemoveAll();
		profile->AccountsList.RemoveAll();
		delete profile;
	}
	ProfilesList.RemoveAll();

	MAPIUninitialize();
	Error=false;
}

bool CMAPIWrapper::RetrieveAll()
{
	bool ret=false;
	LPMAPITABLE pTable = NULL;
	LPPROFADMIN pProfAdmin = NULL; // Pointer to IProfAdmin object
	ULONG ulCount;

	_tprintf(_T("Doing inventory...\r\n"));

	if(FAILED(hResErr = MAPIAdminProfiles(0, &pProfAdmin)))
	{
		_tprintf(_T("MAPIAdminProfiles returned error: %#08x"), hResErr);
		return ret;
	}

	if (FAILED(hResErr = pProfAdmin->GetProfileTable(0, &pTable)))
	{
		_tprintf(_T("GetProfileTable returned error: %#08x"), hResErr);
		pProfAdmin->Release();
		return ret;
	}

	if(SUCCEEDED(hResErr = pTable->GetRowCount(0, &ulCount)))
	{
		if(ulCount > 0)
		{
			SizedSPropTagArray(2, pProfTbl) ={2, {PR_DISPLAY_NAME_A, PR_DEFAULT_PROFILE}};
			LPSRowSet pRows = NULL;
			if(SUCCEEDED(hResErr = HrQueryAllRows(pTable, (LPSPropTagArray)&pProfTbl, NULL, NULL, 0, &pRows)))
			{
				for(ULONG i=0; i< pRows->cRows; i++)
				{
					LPSERVICEADMIN pServiceAdmin=NULL;
					LPMAPITABLE pIMAPITable=NULL;
					LPSRowSet pLPSRowSet = NULL;
					ULONG iRowCount = 0;
					SizedSPropTagArray(1, ptaSvc) = { 1, { PR_SERVICE_UID } }; 
					SRestriction zsres = {0}, res2[2] = {0};
					SPropValue s = {0}, n = {0};
					CMAPIProfile *a = new CMAPIProfile;
					a->isDefault=(bool)pRows->aRow[i].lpProps[1].Value.b;
					a->Name=pRows->aRow[i].lpProps[0].Value.lpszA;
					
					//now huge trick here: to find out what account is default message delivery for particular account
					//we need to query one row GetProviderTable having index 0 
					//we can set it later querying ALL table for columns PR_PROVIDER_UID and using IMsgServiceAdmin::MsgServiceTransportOrder to set it
					hResErr = pProfAdmin->AdminServices((LPTSTR)pRows->aRow[i].lpProps[0].Value.lpszA, NULL, 0, 0,  &pServiceAdmin);
					if(FAILED(hResErr))
					{
						_tprintf(_T("AdminServices returned error: %#08x"), hResErr);
						pTable->Release();
						pProfAdmin->Release();
						return ret;
					}
					hResErr = pServiceAdmin->GetProviderTable(0, &pIMAPITable);
					//pServiceAdmin->DeleteMsgService(
					if(FAILED(hResErr))
					{
						_tprintf(_T("GetProviderTable returned error: %#08x"), hResErr);
						if(pServiceAdmin) pServiceAdmin->Release();
						pTable->Release();
						pProfAdmin->Release();
						return ret;
					}
					res2[0].rt = RES_PROPERTY; 
					res2[0].res.resProperty.relop = RELOP_EQ;
					res2[0].res.resProperty.ulPropTag = PR_PROVIDER_ORDINAL;
					s.ulPropTag=PR_PROVIDER_ORDINAL;
					s.Value.l=0;
					res2[0].res.resProperty.lpProp=&s;

					res2[1].rt = RES_PROPERTY;
					res2[1].res.resProperty.relop = RELOP_EQ;
					res2[1].res.resProperty.ulPropTag = PR_RESOURCE_TYPE;
					n.ulPropTag=PR_RESOURCE_TYPE;
					n.Value.ul=MAPI_TRANSPORT_PROVIDER;
					res2[1].res.resProperty.lpProp=&n;

					zsres.rt = RES_AND;
					zsres.res.resAnd.cRes = 2;
					zsres.res.resAnd.lpRes = &res2[0];

					//
					hResErr = HrQueryAllRows(pIMAPITable, (LPSPropTagArray)&ptaSvc, &zsres, NULL, 0, &pLPSRowSet); 
					if(FAILED(hResErr))
					{
						_tprintf(_T("HrQueryAllRows returned error: %#08x"), hResErr);
					}
					// iRowCount should be equal to 1 
					hResErr = pIMAPITable->GetRowCount(0, &iRowCount); 
					if(iRowCount==1)
						EnumAccounts(a->Name.GetBuffer(), a->AccountsList, a->StoresList, pLPSRowSet->aRow->lpProps[0].Value.bin.lpb, pLPSRowSet->aRow->lpProps[0].Value.bin.cb);
					else if(iRowCount> 1)
					{
						_tprintf(_T("ERROR: RowCount %d > 1 while querying default mapi provider!!!\r\n"), iRowCount);
						EnumAccounts(a->Name.GetBuffer(), a->AccountsList, a->StoresList, NULL, 0);
					}
					else
						EnumAccounts(a->Name.GetBuffer(), a->AccountsList, a->StoresList, NULL, 0);

					ProfilesList.Add(a);

					if(pIMAPITable) pIMAPITable->Release();
					if(pServiceAdmin) pServiceAdmin->Release();
				}
				ret=true;
				FreeProws(pRows);
			}
		}
		else
		{
			_tprintf(_T("No existed mail profiles found!\r\n"));
		}
	}
	else
	{
		_tprintf(_T("GetRowCount on profiles table returned error: %#08x\r\n"), hResErr);
	}
	pTable->Release();
	pProfAdmin->Release();

	return ret;
}

bool CMAPIWrapper::EnumAccounts(LPWSTR lpwszProfile, CMAPIAccount &AccountsList, CMAPIMsgStore &StoresList, LPBYTE DefaultServiceUid, DWORD cbDefaultServiceUid)
{
	HRESULT hRes = S_OK;
	LPMAPISESSION lpSession;
	LPOLKACCOUNTMANAGER lpAcctMgr = NULL;
	
	hRes = MAPILogonEx(0,
		(LPTSTR)lpwszProfile,
		NULL,
		fMapiUnicode | MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE | 
		MAPI_NEW_SESSION | MAPI_NO_MAIL /*| MAPI_LOGON_UI*/,
		&lpSession);

	if(FAILED(hRes))
	{
		_tprintf(_T("Failed MAPILogonEx\r\n"));
		return false;
	}

	//enum stores
	EnumStores(lpSession, StoresList);

	hRes = CoCreateInstance(CLSID_OlkAccountManager, 
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		(LPVOID*)&lpAcctMgr);

	if(SUCCEEDED(hRes) && lpAcctMgr)
	{
		CAccountHelper* pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if(pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = NULL;
			hRes = pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, (LPVOID*)&lpAcctHelper);
			if(SUCCEEDED(hRes) && lpAcctHelper)
			{	
				hRes = lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS);
				if(SUCCEEDED(hRes))
				{
					LPOLKENUM lpAcctEnum = NULL;

					hRes = lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail, &CLSID_OlkMAPIAccount, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum);
					if(SUCCEEDED(hRes) && lpAcctEnum)
					{
						DWORD cAccounts = 0;
						
						hRes = lpAcctEnum->GetCount(&cAccounts);
						if(SUCCEEDED(hRes) && cAccounts)
						{
							hRes = lpAcctEnum->Reset();
							if(SUCCEEDED(hRes))
							{
								DWORD i = 0;
								for (i = 0; i < cAccounts; i++)
								{
									LPUNKNOWN lpUnk = NULL;
									hRes = lpAcctEnum->GetNext(&lpUnk);
									if(SUCCEEDED(hRes) && lpUnk)
									{
										LPOLKACCOUNT lpAccount = NULL;
										hRes = lpUnk->QueryInterface(IID_IOlkAccount, (LPVOID*)&lpAccount);
										if(SUCCEEDED(hRes) && lpAccount)
										{		
											ACCT_VARIANT pProp = {0};
											CMAPIAccount *account = new CMAPIAccount;
											
											//Account ID
											hRes = lpAccount->GetProp(PROP_ACCT_ID, &pProp);
											if(SUCCEEDED(hRes) && pProp.Val.dw)
											{
												account->lAccountID=pProp.Val.dw;
											}
											//Account Name
											hRes = lpAccount->GetProp(PROP_ACCT_NAME, &pProp);
											if(SUCCEEDED(hRes) && pProp.Val.pwsz)
											{	
												account->Name = pProp.Val.pwsz;
											}

											//SERVICE UID
											hRes = lpAccount->GetProp(PROP_MAPI_SERVICE_UID, &pProp);
											if(SUCCEEDED(hRes) && pProp.Val.bin.cb == cbDefaultServiceUid)
											{
												if(memcmp(pProp.Val.bin.pb, DefaultServiceUid, pProp.Val.bin.cb)==0)
													account->isDefault=true;
											}

											//Is Exchange account flag
											hRes = lpAccount->GetProp(PROP_ACCT_IS_EXCH, &pProp);
											if(SUCCEEDED(hRes) && pProp.Val.dw)
											{	
												account->lIsExchange = pProp.Val.dw;
											}
											else
											{
												account->lIsExchange = 0;
											}

											//this will return DEFAULT PR_ENTRYID for account
											hRes = lpAccount->GetProp(PROP_ACCT_DELIVERY_STORE, &pProp);
											if(SUCCEEDED(hRes))
											{	
												account->cbDeliveryStore = pProp.Val.bin.cb;
												account->DeliveryStore = (LPENTRYID)malloc(account->cbDeliveryStore+1);
												memcpy(account->DeliveryStore, pProp.Val.bin.pb, account->cbDeliveryStore);
											}
											else
											{
												_tprintf(_T("PROP_ACCT_DELIVERY_STORE returned error: %#08x\r\n"), hRes);
											}

											hRes = lpAccount->GetProp(PROP_ACCT_DELIVERY_FOLDER, &pProp);
											if(SUCCEEDED(hRes))
											{		
												account->cbDeliveryFolder = pProp.Val.bin.cb;
												account->DeliveryFolder = (LPENTRYID)malloc(account->cbDeliveryFolder+1);
												memcpy(account->DeliveryFolder, pProp.Val.bin.pb, account->cbDeliveryFolder);
											}
											else 
											{
												_tprintf(_T("PROP_ACCT_DELIVERY_FOLDER returned error: %#08x\r\n"), hRes);
											}

											//find out default store for THIS account (current)
											for(int g=0;g<StoresList.GetSize();g++)
											{
												ULONG result;
												CMAPIMsgStore *store=(CMAPIMsgStore *)StoresList.GetAt(g);
												hRes = lpSession->CompareEntryIDs(account->cbDeliveryStore, account->DeliveryStore, store->cbEntryIDstore, store->EntryIDstore,0, &result);
												if(SUCCEEDED(hRes) && result)
												{
													account->MatchedStoreID=g;
												}
											}

											AccountsList.Add(account);
										}
										
										if(lpAccount)
											lpAccount->Release();
										lpAccount = NULL;
									}

									if(lpUnk)
										lpUnk->Release();
									lpUnk = NULL;
								}
							}
						}
					}
					
					if(lpAcctEnum)
						lpAcctEnum->Release();
				}
			}

			if(lpAcctHelper)
				lpAcctHelper->Release();
		}

		if(pMyAcctHelper)
			pMyAcctHelper->Release();
	}	
	
	if(lpAcctMgr)
		lpAcctMgr->Release();


	//not needed maybe lpSession->Release();

	return true;
}

bool CMAPIWrapper::EnumStores(LPMAPISESSION lpSession, CMAPIMsgStore &StoresList)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpTable;
	SizedSPropTagArray(5, ptaTbl) ={5, {PR_DISPLAY_NAME, PR_ENTRYID, PR_DEFAULT_STORE, PR_SERVICE_UID, PR_RESOURCE_FLAGS}};

	hRes = lpSession->GetMsgStoresTable( 0, &lpTable);
    if (SUCCEEDED(hRes)) 
	{
		ULONG rowCount;
		SRestriction zsres = {0};

		//filter stores that can't be selected as default
		zsres.rt = RES_BITMASK;
		zsres.res.resBitMask.ulPropTag = PR_RESOURCE_FLAGS; 
		zsres.res.resBitMask.relBMR = BMR_EQZ; 
		zsres.res.resBitMask.ulMask = STATUS_NO_DEFAULT_STORE;
		//hRes = lpTable->Restrict(&zsres, TBL_BATCH);

		hRes = lpTable->GetRowCount( 0, &rowCount);
        //set query columns display name and entryid
		hRes = lpTable->SetColumns( (LPSPropTagArray)&ptaTbl, 0);
		if (FAILED(hRes)) 
		{
                lpTable->Release();
                return( FALSE);
        }

        hRes = lpTable->SeekRow( BOOKMARK_BEGINNING, 0, NULL);
        if (FAILED(hRes)) 
		{
                lpTable->Release();
                return( FALSE);
        }

		//enum it!
        int cNumRows = 0;
        LPSRowSet lpRow;
        BOOL bResult = TRUE;
		do
		{
			lpRow = NULL;
			hRes = lpTable->QueryRows( 1, 0, &lpRow);
			if(HR_FAILED(hRes)) 
			{
				bResult = FALSE;
				break;
			}

			if(lpRow) 
			{
				cNumRows = lpRow->cRows;
				if (cNumRows)
				{
					CMAPIMsgStore *store = new CMAPIMsgStore;
					//mailstore ID goes here (PR_ENTRYID of the current mailstore)
					LPCTSTR   lpStr = (LPCTSTR) lpRow->aRow[0].lpProps[0].Value.LPSZ;
					LPENTRYID lpEID = (LPENTRYID) lpRow->aRow[0].lpProps[1].Value.bin.lpb;
					ULONG     cbEID = lpRow->aRow[0].lpProps[1].Value.bin.cb;
			
					//PR_SERVICE_UID
					store->cbServiceUID=lpRow->aRow[0].lpProps[3].Value.bin.cb;
					store->ServiceUID=(LPMAPIUID)malloc(lpRow->aRow[0].lpProps[3].Value.bin.cb+1);
					memcpy(store->ServiceUID, lpRow->aRow[0].lpProps[3].Value.bin.lpb, lpRow->aRow[0].lpProps[3].Value.bin.cb);
					//
					store->cbEntryIDstore=cbEID;
					store->EntryIDstore=(LPENTRYID)malloc(cbEID+1);
					memcpy(store->EntryIDstore, lpRow->aRow[0].lpProps[1].Value.bin.lpb, cbEID);
					store->cbStoreName=_tcslen(lpStr);
					if(store->cbStoreName > 0) 
					{
						store->StoreName=(WCHAR*)malloc((store->cbStoreName+1)*sizeof(WCHAR));
						_tcscpy(store->StoreName, lpStr);
					}

					//
					if(lpRow->aRow[0].lpProps[2].Value.b) store->isDefault=true;

					if(!(lpRow->aRow[0].lpProps[4].Value.l&STATUS_NO_DEFAULT_STORE)) store->CanbeDefault=true;

					LPMDB   lpStore;
					hRes = lpSession->OpenMsgStore( NULL, cbEID, lpEID, NULL, MDB_NO_MAIL | MDB_NO_DIALOG, &lpStore);
					if (HR_FAILED(hRes)) 
					{  
					}
					else 
					{
						//default inbox folder ID (PR_ENTRYID of the default inbox folder for current mailstore)
						ULONG cbInboxEID = NULL;
						LPENTRYID lpInboxEID = NULL;
						//PST filePath
						LPSPropValue lpProfilePST=NULL;

						hRes = HrGetOneProp(lpStore, PR_PST_PATH, &lpProfilePST);
						if(SUCCEEDED(hRes))
						{
							store->Path=lpProfilePST->Value.lpszA;
						}

						hRes = lpStore->GetReceiveFolder(NULL, 0, &cbInboxEID, (LPENTRYID *) &lpInboxEID, NULL);
						if(SUCCEEDED(hRes))
						{
							if (cbInboxEID && lpInboxEID)
							{
								store->cbEntryIDInbox=cbInboxEID;
								store->EntryIDInbox=(LPENTRYID)malloc(cbInboxEID+1);
								memcpy(store->EntryIDInbox, lpInboxEID, cbInboxEID);
							}
							MAPIFreeBuffer(lpInboxEID);
						}
						lpStore->Release();
					}
					hRes = S_OK; //force to ok
					StoresList.Add(store);
				}
				FreeProws( lpRow);
			}

		} while ( SUCCEEDED(hRes) && cNumRows && lpRow);
		lpTable->Release();

    }

	return true;
}

bool CMAPIWrapper::SetDefaults(LPWSTR lpwszProfile, LPWSTR accountname, int &StoreIdx, int &ProfileIdx)
{
	HRESULT hRes = S_OK;
	LPMAPISESSION lpSession;
	LPOLKACCOUNTMANAGER lpAcctMgr = NULL;
	int profile_idx=0;

	_tprintf(_T("profile %s accountname %s\r\n"), lpwszProfile, accountname);

	//set default profile
	LPPROFADMIN pProfAdmin = NULL; // Pointer to IProfAdmin object
	if(FAILED(hResErr = MAPIAdminProfiles(0, &pProfAdmin)))
	{
		_tprintf(_T("MAPIAdminProfiles returned error: %#08x"), hResErr);
		return false;
	}

	hResErr = pProfAdmin->SetDefaultProfile((LPTSTR)lpwszProfile, 0);
	if(FAILED(hRes))
	{
		_tprintf(_T("Failed SetDefaultProfile\r\n"));
		return false;
	}
	else
	{
		_tprintf(_T("Done set default profile\r\n"));
	}
	pProfAdmin->Release();


	hRes = MAPILogonEx(0,
		(LPTSTR)lpwszProfile,
		NULL,
		fMapiUnicode | MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE | 
		MAPI_NEW_SESSION | MAPI_NO_MAIL /*| MAPI_LOGON_UI*/,
		&lpSession);
	if(FAILED(hRes))
	{
		_tprintf(_T("Failed MAPILogonEx\r\n"));
		return false;
	}

	//set default delivery (first obtain ENTRYID from our known inventory data)
	CMAPIProfile *profile=(CMAPIProfile *)ProfilesList.GetAt(ProfileIdx);
	CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(StoreIdx);
	hRes=lpSession->SetDefaultStore(0, store->cbEntryIDstore, store->EntryIDstore);
	if(FAILED(hRes))
	{
		_tprintf(_T("Failed SetDefaultStore\r\n"));
	}
	else
	{
		_tprintf(_T("Done set default store\r\n"));
	}

	hRes = CoCreateInstance(CLSID_OlkAccountManager, 
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		(LPVOID*)&lpAcctMgr);

	if(SUCCEEDED(hRes) && lpAcctMgr)
	{
		CAccountHelper* pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if(pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = NULL;
			hRes = pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, (LPVOID*)&lpAcctHelper);
			if(SUCCEEDED(hRes) && lpAcctHelper)
			{	
				hRes = lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS);
				if(SUCCEEDED(hRes))
				{
					LPOLKENUM lpAcctEnum = NULL;

					hRes = lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail, &CLSID_OlkMAPIAccount, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum);
					if(SUCCEEDED(hRes) && lpAcctEnum)
					{
						DWORD cAccounts = 0;
						
						hRes = lpAcctEnum->GetCount(&cAccounts);
						if(SUCCEEDED(hRes) && cAccounts)
						{
							hRes = lpAcctEnum->Reset();
							if(SUCCEEDED(hRes))
							{
								DWORD i = 0;
								for (i = 0; i < cAccounts; i++)
								{
									LPUNKNOWN lpUnk = NULL;
									hRes = lpAcctEnum->GetNext(&lpUnk);
									if(SUCCEEDED(hRes) && lpUnk)
									{
										LPOLKACCOUNT lpAccount = NULL;
										hRes = lpUnk->QueryInterface(IID_IOlkAccount, (LPVOID*)&lpAccount);
										if(SUCCEEDED(hRes) && lpAccount)
										{		
											ACCT_VARIANT pProp = {0};
											
											//Account Name
											hRes = lpAccount->GetProp(PROP_ACCT_NAME, &pProp);
											if(SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												//check: is current account equal requested account?
												if(_tcscmp(pProp.Val.pwsz, accountname)==0)
												{
													ACCT_VARIANT Prop1 = {0}, Prop2 = {0};
													_tprintf(_T("starting set delivery store\r\n"));
													Prop1.dwType=PT_BINARY;
													Prop1.Val.bin.cb=store->cbEntryIDstore;
													Prop1.Val.bin.pb=(BYTE *)store->EntryIDstore;
													hRes=lpAccount->SetProp(PROP_ACCT_DELIVERY_STORE, &Prop1);
													if(FAILED(hRes))
													{
														_tprintf(_T("Account Set PROP_ACCT_DELIVERY_STORE returned error: %#08x\r\n"), hRes);
													}
													Prop2.dwType=PT_BINARY;
													Prop2.Val.bin.cb=store->cbEntryIDInbox;
													Prop2.Val.bin.pb=(BYTE *)store->EntryIDInbox;
													hRes=lpAccount->SetProp(PROP_ACCT_DELIVERY_FOLDER, &Prop2);
													if(FAILED(hRes))
													{
														_tprintf(_T("Account Set PROP_ACCT_DELIVERY_FOLDER returned error: %#08x\r\n"), hRes);
													}
													hRes=lpAccount->SaveChanges(OLK_ACCOUNT_NO_FLAGS);
													if(FAILED(hRes))
													{
														_tprintf(_T("Account SaveChanges returned error: %#08x\r\n"), hRes);
													}
													else 
													{
														_tprintf(_T("Done set default delivery\r\n"));
													}
												}
											}
										}
										
										if(lpAccount)
											lpAccount->Release();
										lpAccount = NULL;
									}

									if(lpUnk)
										lpUnk->Release();
									lpUnk = NULL;
								}
							}
						}
					}
					
					if(lpAcctEnum)
						lpAcctEnum->Release();
				}
			}

			if(lpAcctHelper)
				lpAcctHelper->Release();
		}

		if(pMyAcctHelper)
			pMyAcctHelper->Release();
	}	
	
	if(lpAcctMgr)
		lpAcctMgr->Release();


	//not needed maybe lpSession->Release();

	return true;
}

bool CMAPIWrapper::RemoveStore(LPWSTR lpwszProfile, int &StoreIdx, int &ProfileIdx)
{
	bool ret_code=false;
	HRESULT hRes = S_OK;
	LPMAPISESSION lpSession;
	LPSERVICEADMIN pServiceAdmin=NULL;

	_tprintf(_T("profile %s store index %d\r\n"), lpwszProfile, StoreIdx);

	hRes = MAPILogonEx(0,
		(LPTSTR)lpwszProfile,
		NULL,
		fMapiUnicode | MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE | 
		MAPI_NEW_SESSION | MAPI_NO_MAIL,
		&lpSession);

	if(FAILED(hRes))
	{
		_tprintf(_T("Failed MAPILogonEx\r\n"));
		return ret_code;
	}

	hResErr = lpSession->AdminServices(0, &pServiceAdmin);
	if(FAILED(hResErr))
	{
		_tprintf(_T("AdminServices returned error: %#08x\r\n"), hResErr);
		return ret_code;
	}

	CMAPIProfile *profile=(CMAPIProfile *)ProfilesList.GetAt(ProfileIdx);
	CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(StoreIdx);
	hResErr = pServiceAdmin->DeleteMsgService((LPMAPIUID)store->ServiceUID);
	if(FAILED(hResErr))
	{
		_tprintf(_T("DeleteMsgService returned error: %#08x\r\n"), hResErr);
	}
	else
	{
		_tprintf(_T("Done removing!\r\n"));
		ret_code=true;
	}

	if(pServiceAdmin) pServiceAdmin->Release();
	//not needed maybe lpSession->Release();

	return ret_code;
}
