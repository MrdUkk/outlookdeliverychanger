/*
   This is my derivative work based on article:
   "Using Account Management API (IOlkAccountManger) to List Outlook Email Accounts"
   by Ashutosh Bhawasinka,  28 Aug 2008
   published first on www.codeproject.com
   
   (c) dUkk 2011-2014
*/
#pragma once


class CMAPIMsgStore: public CObArray 
{
public:
	CMAPIMsgStore() { CanbeDefault=false; isDefault=false; StoreName=NULL; cbEntryIDstore=0; EntryIDstore=NULL; cbEntryIDInbox=0; EntryIDInbox=NULL; ServiceUID=NULL; cbServiceUID=0; };
	//~CMAPIMsgStore() { if(StoreName) free(StoreName); if(EntryIDstore) free(EntryIDstore); if(EntryIDInbox) free(EntryIDInbox); if(ServiceUID) free(ServiceUID); };
	LPENTRYID EntryIDstore;  //message store ENTRYID
	ULONG cbEntryIDstore;    //its sizeof
	LPENTRYID EntryIDInbox;  //message store ENTRYID
	ULONG cbEntryIDInbox;    //its sizeof
	LPMAPIUID ServiceUID;    //message store SERVICEUID
	ULONG cbServiceUID;		 //its sizeof
	WCHAR *StoreName;        //message store human readable name
	int cbStoreName;		 //sizeof message store human name
	CString Path;
	bool isDefault;
	bool CanbeDefault;
};

class CMAPIAccount: public CObArray
{
public:
	CMAPIAccount() { isDefault=false; MatchedStoreID=-1; cbDeliveryStore=0; DeliveryStore=NULL; DeliveryFolder=NULL; cbDeliveryFolder=0; };
	//~CMAPIAccount() { if(DeliveryStore) free(DeliveryStore); if(DeliveryFolder) free(DeliveryFolder); };
	CString Name;   //human readable account name
	LPENTRYID DeliveryStore; //ENTRYID of default delivery message store
	ULONG cbDeliveryStore;   //its sizeof ENTRYID
	LPENTRYID DeliveryFolder; //ENTRYID of Inbox folder in default delivery message store
	ULONG cbDeliveryFolder;  //its sizeof ENTRYID
	long lAccountID;
	long lIsExchange;
	bool isDefault;
	int MatchedStoreID;   //index of element matched with some EntryStoreID in CMAPIMsgStore objects array
};

class CMAPIProfile: public CObArray
{
public:
	CMAPIProfile() { isDefault=false; };
	CString Name;
	bool isDefault;
	CMAPIAccount AccountsList;
	CMAPIMsgStore StoresList;
};

class CMAPIWrapper {
public:
	CMAPIProfile ProfilesList; //our main huge data holder :)
	bool Error; 	//global flag indicating problems (errors)
	HRESULT hResErr; //last operation error

	CMAPIWrapper();
	~CMAPIWrapper();

	//
	bool RetrieveAll();
	bool SetDefaults(LPWSTR lpwszProfile, LPWSTR accountname, int &StoreIdx, int &ProfileIdx);
	bool RemoveStore(LPWSTR lpwszProfile, int &StoreIdx, int &ProfileIdx);

private:
	bool EnumAccounts(LPWSTR lpwszProfile, CMAPIAccount &AccountsList, CMAPIMsgStore &StoresList, LPBYTE DefaultServiceUid, DWORD cbDefaultServiceUid);
	bool EnumStores(LPMAPISESSION lpSession, CMAPIMsgStore &StoresList);
};