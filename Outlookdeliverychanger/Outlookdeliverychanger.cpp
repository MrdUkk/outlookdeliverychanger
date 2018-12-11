/*
	Main entry point
	
	work was done after investigating and research using those articles as is	
	http://blogs.msdn.com/b/jmazner/archive/2006/10/30/setting-autoarchive-properties-on-a-folder-hierarchy-in-outlook-2007.aspx

	http://www.pcreview.co.uk/forums/script-enable-autoarchive-outlook-2003-a-t3238865.html

	http://support.microsoft.com/kb/194955

	http://groups.google.com/group/microsoft.public.outlook.program_vba/browse_thread/thread/ce7ded7a53fe95fa/8709b7c0a7f63032?lnk=gst&q=auto+archive+eric+legault&rnum=5#8709b7c0a7f63032

	http://support.microsoft.com/kb/198479

	http://stackoverflow.com/questions/624377/detecting-autoarchive-settings-store-in-outlook-2007

	https://www.codeproject.com/articles/27494/using-account-management-api-iolkaccountmanger-to
	
	To compile you will need:
	+MS Office 2010 MAPI Headers
	+MAPIStublibrary (old https://archive.codeplex.com/?p=mapistublibrary new https://github.com/stephenegriffin/MAPIStubLibrary)
	
   (c) dUkk 2011-2014
*/

#include "stdafx.h"
#include "Outlookdeliverychanger.h"
#include "AccountHelper.h"
#include <MAPIUTIL.H>
#include "tinyxml2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CWinApp theApp;

using namespace std;
using namespace tinyxml2;

void printusage();
bool isOutlookPresent(void);

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int ret_code=0;

	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		_tprintf(_T("Fatal Error: MFC initialization failed\n"));
		return -1;
	}

#ifdef _WIN64
	_tprintf(_T("MS Office Outlook delivery changer 64bit v1.6 by dUkk (c) 2014\r\n"));
#else
	_tprintf(_T("MS Office Outlook delivery changer v1.6 by dUkk (c) 2011-2014\r\n"));
#endif
	
	if(argc < 2) 
	{
		printusage();
		return -2;
	}

	if(!isOutlookPresent())
	{
		return -3;
	}


	CMAPIWrapper my;
	if(!my.RetrieveAll())
	{
		return -3;
	}

	if(_tcsicmp(argv[1],_T("/list"))==0)
	{
		int x=my.ProfilesList.GetSize();
		_tprintf(_T("Total profiles: %d\r\n"), x);

		for(int a=0;a<x;a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			int acc_c=profile->AccountsList.GetSize();
			int msgstores_c=profile->StoresList.GetSize();
			_tprintf(_T("Profile name: [%s] , has %d accounts , %d message stores, default: %d\r\n"), profile->Name, acc_c, msgstores_c, profile->isDefault);
			//list stores
			for(int c=0;c<msgstores_c;c++)
			{
				CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);

				_tprintf(_T("  Store name: [%s]"), store->StoreName);
				if(!store->Path.IsEmpty())
				{
					_tprintf(_T(" File path: [%s]"), store->Path);
				}
				else
				{
					_tprintf(_T(" Server-type"));
				}
				if(store->isDefault)
				{
					_tprintf(_T(" , *DEFAULT STORE FOR PROFILE*"));
				}
				_tprintf(_T("\r\n"));
			}
			
			//list accounts with its stores delivery
			for(int b=0;b<acc_c;b++)
			{
				CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(b);
				_tprintf(_T(" Account name: [%s] MS Exchange: %d, default: %d\r\n"), account->Name, account->lIsExchange, account->isDefault);
				if(account->MatchedStoreID != -1)
				{
					CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(account->MatchedStoreID);
					_tprintf(_T("  will deliver new messages to store: [%s]"), store->StoreName);
				}
				_tprintf(_T("\r\n"));
			}
		}
	}
	else if(_tcsicmp(argv[1],_T("/listxml"))==0 && argc > 2)
	{
		TiXMLDocument mdoc;
        XMLNode *Root;
		mdoc.Parse(_T("<?xml version=\"1.0\" encoding=\"utf-16\" ?><odc></odc>"));
        Root = mdoc.RootElement();

		int x=my.ProfilesList.GetSize();
		for(int a=0;a<x;a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			int acc_c=profile->AccountsList.GetSize();
			int msgstores_c=profile->StoresList.GetSize();

			XMLElement *item1 = mdoc.NewElement(_T("profile"));
			item1->SetAttribute(_T("name"), profile->Name);
			item1->SetAttribute(_T("isdefault"), profile->isDefault);
			XMLNode *Subnode=Root->InsertEndChild(item1);
			//list accounts with its stores delivery
			for(int b=0;b<acc_c;b++)
			{
				CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(b);
				XMLElement *item2 = mdoc.NewElement(_T("account"));
				item2->SetAttribute(_T("name"), account->Name);
				item2->SetAttribute(_T("isexchange"), account->lIsExchange);
				item2->SetAttribute(_T("isdefault"), account->isDefault);
				if(account->MatchedStoreID != -1)
				{
					CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(account->MatchedStoreID);
					item2->SetAttribute(_T("deliverto"), store->StoreName);
				}
				Subnode->InsertEndChild(item2);
			}

			//list stores
			for(int c=0;c<msgstores_c;c++)
			{
				CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
				XMLElement *item2 = mdoc.NewElement(_T("store"));
				item2->SetAttribute(_T("name"), store->StoreName);

				if(!store->Path.IsEmpty())
				{
					item2->SetAttribute(_T("type"), store->Path);
				}
				else
				{
					item2->SetAttribute(_T("type"), _T("Server"));
				}
				item2->SetAttribute(_T("isdefault"), store->isDefault);
				Subnode->InsertEndChild(item2);
			}			
		}
		mdoc.SaveFile(argv[2]);
	}
	else if(_tcsicmp(argv[1],_T("/exportdef"))==0 && argc > 2)
	{
		TiXMLDocument mdoc;
        XMLNode *Root;
		mdoc.Parse(_T("<?xml version=\"1.0\" encoding=\"utf-16\" ?><odc></odc>"));
		Root = mdoc.RootElement();

		int x=my.ProfilesList.GetSize();
		for(int a=0;a<x;a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			int acc_c=profile->AccountsList.GetSize();
			if(profile->isDefault)
			{
				XMLElement *item1 = mdoc.NewElement(_T("profiledefault"));
				item1->SetAttribute(_T("name"), profile->Name);
				Root->InsertEndChild(item1);

				for(int b=0;b<acc_c;b++)
				{
					CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(b);
					//valid is for exchange DEFAULT account
					if(account->lIsExchange && account->isDefault)
					{
						XMLElement *item2 = mdoc.NewElement(_T("accountdefault"));
						item2->SetAttribute(_T("name"), account->Name);
						Root->InsertEndChild(item2);

						if(account->MatchedStoreID != -1)
						{
							CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(account->MatchedStoreID);
							XMLElement *item3 = mdoc.NewElement(_T("storedefault"));
							item3->SetAttribute(_T("name"), store->StoreName);
							if(!store->Path.IsEmpty())
							{
								item3->SetAttribute(_T("filepath"), store->Path);
							}
							Root->InsertEndChild(item3);
						}

						int msgstores_c=profile->StoresList.GetSize();
						for(int c=0;c<msgstores_c;c++)
						{
							CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
							if(store->isDefault)
							{
								XMLElement *item3 = mdoc.NewElement(_T("deliverydefault"));
								item3->SetAttribute(_T("name"), store->StoreName);
								if(!store->Path.IsEmpty())
								{
									item3->SetAttribute(_T("filepath"), store->Path);
								}
								Root->InsertEndChild(item3);
							}
						}

					}
				}
			}
		}
		mdoc.SaveFile(argv[2]);
	}
	else if(_tcsicmp(argv[1],_T("/importdef"))==0 && argc > 2)
	{
		TiXMLDocument mdoc;
		XMLNode *Root;
		XMLElement *Element;
		TCHAR profilename[2048], accountname[2048], defaultstore[2048] = {0}, defaultdelivery[2048] = {0};
		bool defaultstoreFile, defaultdeliveryFile;
		int store_idx=0, profile_idx=0;

		if (mdoc.LoadFile(argv[2]) != XML_SUCCESS)
		{
			_tprintf(_T("ERROR opening/parsing XML file: %s\r\n"), argv[2]);
			return -4;
		}

        Root = mdoc.RootElement();

		Element = Root->FirstChildElement(_T("profiledefault"));
		if (Element)
		{
			if (Element->Attribute(_T("name"))) wcscpy_s(profilename, Element->Attribute(_T("name")));
		}

		Element = Root->FirstChildElement(_T("accountdefault"));
		if (Element)
		{
			if (Element->Attribute(_T("name"))) wcscpy_s(accountname, Element->Attribute(_T("name")));
		}

		Element = Root->FirstChildElement(_T("storedefault"));
		if (Element)
		{
			if(Element->Attribute(_T("filepath")))
			{
				wcscpy_s(defaultstore, Element->Attribute(_T("filepath")));
				defaultstoreFile=true;
			}
			else if(Element->Attribute(_T("name")))
			{
				wcscpy_s(defaultstore, Element->Attribute(_T("name")));
				defaultstoreFile=false;
			}
		}

		Element = Root->FirstChildElement(_T("deliverydefault"));
		if (Element)
		{
			if(Element->Attribute(_T("filepath"))) 
			{
				wcscpy_s(defaultdelivery, Element->Attribute(_T("filepath")));
				defaultdeliveryFile=true;
			}
			else if(Element->Attribute(_T("name")))
			{
				wcscpy_s(defaultdelivery, Element->Attribute(_T("name")));
				defaultdeliveryFile=false;
			}
		}

		//defaults obtained now try find its indexes in current configuration
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->Name.CompareNoCase(profilename)==0)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search default account
				for(int b=0;b<profile->AccountsList.GetSize();b++)
				{
					CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(b);
					if(account->Name.CompareNoCase(accountname)==0)
					{
						_tcscpy(accountname, account->Name);
						//search mailbox store
						for(int c=0;c<profile->StoresList.GetSize();c++)
						{
							CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
							store_idx=c;
							//probably this is MAILBOX (because it doesn't have file path)
							if(defaultstoreFile==false || defaultdeliveryFile==false)
							{
								if(_tcsicmp(store->StoreName, defaultstore)==0 || _tcsicmp(store->StoreName, defaultdelivery)==0)
								{
									if(my.SetDefaults((LPWSTR)&profilename, (LPWSTR)&accountname, store_idx, profile_idx)) ret_code=0;
									else ret_code=-31;
								}
							}
							else if(defaultstoreFile==true || defaultdeliveryFile==true)
							{
								if(store->Path.CompareNoCase(defaultstore)==0 || store->Path.CompareNoCase(defaultdelivery)==0)
								{
									if(my.SetDefaults((LPWSTR)&profilename, (LPWSTR)&accountname, store_idx, profile_idx)) ret_code=0;
									else ret_code=-31;
								}
							}
						}
					}
				}
			}
		}

	}
	else if(_tcsicmp(argv[1],_T("/setdefmailbox"))==0)
	{
		TCHAR profilename[2048], accountname[2048];
		int store_idx=0, profile_idx=0;

		ret_code=-30;
		//search default profile
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->isDefault)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search default account
				for(int b=0;b<profile->AccountsList.GetSize();b++)
				{
					CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(b);
					if(account->isDefault)
					{
						_tcscpy(accountname, account->Name);
						//find mailbox-type store
						for(int c=0;c<profile->StoresList.GetSize();c++)
						{
							CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
							//probably this is MAILBOX (because it doesn't have file path)
							if(store->Path.IsEmpty() && store->CanbeDefault)
							{
								store_idx=c;
								if(my.SetDefaults((LPWSTR)&profilename, (LPWSTR)&accountname, store_idx, profile_idx)) ret_code=0;
								else ret_code=-31;
							}
						}
					}
				}
			}
		}
	}
	else if(_tcsicmp(argv[1],_T("/setdefbypath"))==0 && argc > 2)
	{
		TCHAR profilename[2048], accountname[2048];
		int store_idx=0, profile_idx=0;

		ret_code=-40;
		//search default profile
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->isDefault)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search default account
				for(int b=0;b<profile->AccountsList.GetSize();b++)
				{
					CMAPIAccount *account=(CMAPIAccount *)profile->AccountsList.GetAt(b);
					if(account->isDefault)
					{
						_tcscpy(accountname, account->Name);
						//search specified store by its path
						for(int c=0;c<profile->StoresList.GetSize();c++)
						{
							CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
							if(store->Path.CompareNoCase(argv[2])==0)
							{
								store_idx=c;
								if(my.SetDefaults((LPWSTR)&profilename, (LPWSTR)&accountname, store_idx, profile_idx)) ret_code=0;
								else ret_code=-41;
							}
						}
					}
				}
			}
		}
	}
	else if(_tcsicmp(argv[1],_T("/delstorebypath"))==0 && argc > 2)
	{
		TCHAR profilename[2048] = {0};
		int store_idx=0, profile_idx=0;

		ret_code=-50;
		//search default profile
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->isDefault)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search specified store
				for(int c=0;c<profile->StoresList.GetSize();c++)
				{
					CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
					if(store->Path.CompareNoCase(argv[2])==0)
					{
						store_idx=c;
						if(my.RemoveStore((LPWSTR)&profilename, store_idx, profile_idx)) ret_code=0;
						else ret_code=-51;
					}
				}
			}
		}
	}
	else if(_tcsicmp(argv[1],_T("/dellocalstores"))==0)
	{
		TCHAR profilename[2048] = {0};
		int store_idx=0, profile_idx=0;

		ret_code=-60;
		//search default profile
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->isDefault)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search specified store
				for(int c=0;c<profile->StoresList.GetSize();c++)
				{
					CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
					if(!store->Path.IsEmpty() && store->Path.Find(_T(":\\"), 0) != -1)
					{
						store_idx=c;
						if(my.RemoveStore((LPWSTR)&profilename, store_idx, profile_idx)) ret_code=0;
						else ret_code=-61;
					}
				}
			}
		}
	}
	else if(_tcsicmp(argv[1],_T("/delnondefstores"))==0)
	{
		TCHAR profilename[2048] = {0};
		int store_idx=0, profile_idx=0;

		ret_code=-70;
		//search default profile
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->isDefault)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search all non default stores and disconnect it from profile
				for(int c=0;c<profile->StoresList.GetSize();c++)
				{
					CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
					if(!store->isDefault && !store->Path.IsEmpty() && store->CanbeDefault)
					{
						store_idx=c;
						if(my.RemoveStore((LPWSTR)&profilename, store_idx, profile_idx)) ret_code=0;
						else ret_code=-71;
					}
				}
			}
		}
	}
	//secret special case for some company (disconnect psts on mounted disk X: fileserver)
	else if(_tcsicmp(argv[1],_T("/hiddenopt"))==0)
	{
		TCHAR profilename[2048] = {0};
		int store_idx=0, profile_idx=0;

		ret_code=-60;
		//search default profile
		for(int a=0;a<my.ProfilesList.GetSize();a++)
		{
			CMAPIProfile *profile=(CMAPIProfile *)my.ProfilesList.GetAt(a);
			if(profile->isDefault)
			{
				_tcscpy(profilename, profile->Name);
				profile_idx=a;
				//search specified store
				for(int c=0;c<profile->StoresList.GetSize();c++)
				{
					CMAPIMsgStore *store=(CMAPIMsgStore *)profile->StoresList.GetAt(c);
					if(!store->Path.IsEmpty() && store->Path.Find(_T("X:\\"), 0) != -1)
					{
						store_idx=c;
						if(my.RemoveStore((LPWSTR)&profilename, store_idx, profile_idx)) ret_code=0;
						else ret_code=-61;
					}
				}
			}
		}
	}
	else
	{
		_tprintf(_T("unknown or insufficient parameters!\r\n\r\n"));
		printusage();
		return -2;
	}

	_tprintf(_T("\r\n\r\nfinished with = %d \r\n"), ret_code);

	return ret_code;
}


void printusage()
{
	_tprintf(_T("\r\n"));
	_tprintf(_T("tested on MS Outlook 2003/2007/2010/2013 versions\r\n"));
	_tprintf(_T("usage: outlookdeliverychanger.exe <function> <parameter>\r\n"));
	_tprintf(_T("       where <function> is one of the:\r\n"));
	_tprintf(_T("       /list            - print on console all mail profiles,accounts,stores for current user\r\n"));
	_tprintf(_T("       /listxml         - export to XML file specified in <parameter> all mail profiles,accounts,stores for current user\r\n"));
	_tprintf(_T("       /exportdef       - export default delivery to XML file specified in <parameter>\r\n"));
	_tprintf(_T("       /importdef       - import default delivery from XML file specified in <parameter>\r\n"));
	_tprintf(_T("       /setdefmailbox   - set default delivery to mailbox\r\n"));
	_tprintf(_T("       /setdefbypath    - set default delivery to PST store path in <parameter>\r\n"));
	_tprintf(_T("       /delstorebypath  - remove store from profile by full path in <parameter>\r\n"));
	_tprintf(_T("       /dellocalstores  - remove all LOCAL stores from profile (not on exchange server)\r\n"));
	_tprintf(_T("       /delnondefstores - remove all NON DEFAULT delivery (files) stores from profile\r\n"));
	_tprintf(_T("\r\n"));
	_tprintf(_T("       NOTE: all write operations will be performed on CURRENT (default) profile and account!\r\n"));
	_tprintf(_T("\r\n\r\n"));
}


bool isOutlookPresent(void)
{
	HRESULT hr = S_OK;
	DWORD dwSize;
	DWORD dwType;
	HKEY hk1 = NULL;
	HKEY hkMAPI2 = NULL;
	WCHAR mainkeyname[1024];
	WCHAR rgwchMSILCID[32];
	const WCHAR *Versions[] = {_T("15.0"),_T("14.0"),_T("13.0"), _T("12.0"), _T("11.0"), _T("10.0"), _T("9.0"), _T("8.0")};

	for(int a=0;a< 5;a++) 
	{
		_swprintf(mainkeyname, _T("Software\\Microsoft\\Office\\%s\\Outlook"), Versions[a]);
		if(RegOpenKeyExW(HKEY_LOCAL_MACHINE, mainkeyname, 0, KEY_READ, &hk1) == ERROR_SUCCESS)
		{
			_tprintf(_T("Outlook %s found"), Versions[a]);
			dwSize = sizeof(rgwchMSILCID);
			hr = RegQueryValueExW(hk1, _T("Bitness"), 0, &dwType, (LPBYTE)rgwchMSILCID, &dwSize);
			if ((ERROR_SUCCESS == hr) && (REG_SZ == dwType))
			{
				RegCloseKey(hk1);
				if(wcscmp(rgwchMSILCID, _T("x86")) ==0)
				{
					_tprintf(_T(" 32bit version\r\n"));
#ifdef _WIN64
					_tprintf(_T(" use another architecture!\r\n"));
					return false;
#else
					return true;
#endif
				}
				else {
					_tprintf(_T(" 64bit version\r\n"));
#ifdef _WIN64
					return true;
#else
					_tprintf(_T(" use another architecture!\r\n"));
					return false;
#endif
				}
			}
			RegCloseKey(hk1);
			_tprintf(_T("\r\n"));
#ifdef _WIN64
			return false;
#else
			return true;
#endif
		}
	}

	_tprintf(_T("no MAPI detected for this architecture\r\n"));
	return false;
}