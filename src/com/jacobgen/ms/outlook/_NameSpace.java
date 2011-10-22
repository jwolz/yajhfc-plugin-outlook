/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _NameSpace extends Dispatch {

	public static final String componentName = "Outlook._NameSpace";

	public _NameSpace() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _NameSpace(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _NameSpace(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Application
	 */
	public _Application getApplication() {
		return new _Application(Dispatch.get(this, "Application").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, "Class").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, "Session").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getParent() {
		return Dispatch.get(this, "Parent");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Recipient
	 */
	public Recipient getCurrentUser() {
		return new Recipient(Dispatch.get(this, "CurrentUser").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Folders
	 */
	public _Folders getFolders() {
		return new _Folders(Dispatch.get(this, "Folders").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getType() {
		return Dispatch.get(this, "Type").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressLists
	 */
	public AddressLists getAddressLists() {
		return new AddressLists(Dispatch.get(this, "AddressLists").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param recipientName an input-parameter of type String
	 * @return the result is of type Recipient
	 */
	public Recipient createRecipient(String recipientName) {
		return new Recipient(Dispatch.call(this, "CreateRecipient", recipientName).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folderType an input-parameter of type int
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getDefaultFolder(int folderType) {
		return new MAPIFolder(Dispatch.call(this, "GetDefaultFolder", new Variant(folderType)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param entryIDFolder an input-parameter of type String
	 * @param entryIDStore an input-parameter of type Variant
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getFolderFromID(String entryIDFolder, Variant entryIDStore) {
		return new MAPIFolder(Dispatch.call(this, "GetFolderFromID", entryIDFolder, entryIDStore).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param entryIDFolder an input-parameter of type String
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getFolderFromID(String entryIDFolder) {
		return new MAPIFolder(Dispatch.call(this, "GetFolderFromID", entryIDFolder).toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param entryIDItem an input-parameter of type String
	 * @param entryIDStore an input-parameter of type Variant
	 * @return the result is of type Object
	 */
	public Object getItemFromID(String entryIDItem, Variant entryIDStore) {
		return Dispatch.call(this, "GetItemFromID", entryIDItem, entryIDStore);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param entryIDItem an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object getItemFromID(String entryIDItem) {
		return Dispatch.call(this, "GetItemFromID", entryIDItem);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param entryID an input-parameter of type String
	 * @return the result is of type Recipient
	 */
	public Recipient getRecipientFromID(String entryID) {
		return new Recipient(Dispatch.call(this, "GetRecipientFromID", entryID).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param recipient an input-parameter of type Recipient
	 * @param folderType an input-parameter of type int
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getSharedDefaultFolder(Recipient recipient, int folderType) {
		return new MAPIFolder(Dispatch.call(this, "GetSharedDefaultFolder", recipient, new Variant(folderType)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void logoff() {
		Dispatch.call(this, "Logoff");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param profile an input-parameter of type Variant
	 * @param password an input-parameter of type Variant
	 * @param showDialog an input-parameter of type Variant
	 * @param newSession an input-parameter of type Variant
	 */
	public void logon(Variant profile, Variant password, Variant showDialog, Variant newSession) {
		Dispatch.call(this, "Logon", profile, password, showDialog, newSession);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param profile an input-parameter of type Variant
	 * @param password an input-parameter of type Variant
	 * @param showDialog an input-parameter of type Variant
	 */
	public void logon(Variant profile, Variant password, Variant showDialog) {
		Dispatch.call(this, "Logon", profile, password, showDialog);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param profile an input-parameter of type Variant
	 * @param password an input-parameter of type Variant
	 */
	public void logon(Variant profile, Variant password) {
		Dispatch.call(this, "Logon", profile, password);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param profile an input-parameter of type Variant
	 */
	public void logon(Variant profile) {
		Dispatch.call(this, "Logon", profile);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void logon() {
		Dispatch.call(this, "Logon");
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder pickFolder() {
		return new MAPIFolder(Dispatch.call(this, "PickFolder").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void refreshRemoteHeaders() {
		Dispatch.call(this, "RefreshRemoteHeaders");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SyncObjects
	 */
	public SyncObjects getSyncObjects() {
		return new SyncObjects(Dispatch.get(this, "SyncObjects").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param store an input-parameter of type Variant
	 */
	public void addStore(Variant store) {
		Dispatch.call(this, "AddStore", store);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folder an input-parameter of type MAPIFolder
	 */
	public void removeStore(MAPIFolder folder) {
		Dispatch.call(this, "RemoveStore", folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getOffline() {
		return Dispatch.get(this, "Offline").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param contactItem an input-parameter of type Variant
	 */
	public void dial(Variant contactItem) {
		Dispatch.call(this, "Dial", contactItem);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void dial() {
		Dispatch.call(this, "Dial");
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param contactItem an input-parameter of type Variant
//	 */
//	public void dial(Variant contactItem) {
//		Dispatch.call(this, "Dial", contactItem);
//
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getMAPIOBJECT() {
		return Dispatch.get(this, "MAPIOBJECT");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getExchangeConnectionMode() {
		return Dispatch.get(this, "ExchangeConnectionMode").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param store an input-parameter of type Variant
	 * @param type an input-parameter of type int
	 */
	public void addStoreEx(Variant store, int type) {
		Dispatch.call(this, "AddStoreEx", store, new Variant(type));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Accounts
	 */
	public Accounts getAccounts() {
		return new Accounts(Dispatch.get(this, "Accounts").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCurrentProfileName() {
		return Dispatch.get(this, "CurrentProfileName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Stores
	 */
	public Stores getStores() {
		return new Stores(Dispatch.get(this, "Stores").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SelectNamesDialog
	 */
	public SelectNamesDialog getSelectNamesDialog() {
		return new SelectNamesDialog(Dispatch.call(this, "GetSelectNamesDialog").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showProgressDialog an input-parameter of type boolean
	 */
	public void sendAndReceive(boolean showProgressDialog) {
		Dispatch.call(this, "SendAndReceive", new Variant(showProgressDialog));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Store
	 */
	public Store getDefaultStore() {
		return new Store(Dispatch.get(this, "DefaultStore").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param iD an input-parameter of type String
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getAddressEntryFromID(String iD) {
		return new AddressEntry(Dispatch.call(this, "GetAddressEntryFromID", iD).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressList
	 */
	public AddressList getGlobalAddressList() {
		return new AddressList(Dispatch.call(this, "GetGlobalAddressList").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param iD an input-parameter of type String
	 * @return the result is of type Store
	 */
	public Store getStoreFromID(String iD) {
		return new Store(Dispatch.call(this, "GetStoreFromID", iD).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Categories
	 */
	public Categories getCategories() {
		return new Categories(Dispatch.get(this, "Categories").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @param name an input-parameter of type Variant
	 * @param downloadAttachments an input-parameter of type Variant
	 * @param useTTL an input-parameter of type Variant
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder openSharedFolder(String path, Variant name, Variant downloadAttachments, Variant useTTL) {
		return new MAPIFolder(Dispatch.call(this, "OpenSharedFolder", path, name, downloadAttachments, useTTL).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @param name an input-parameter of type Variant
	 * @param downloadAttachments an input-parameter of type Variant
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder openSharedFolder(String path, Variant name, Variant downloadAttachments) {
		return new MAPIFolder(Dispatch.call(this, "OpenSharedFolder", path, name, downloadAttachments).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @param name an input-parameter of type Variant
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder openSharedFolder(String path, Variant name) {
		return new MAPIFolder(Dispatch.call(this, "OpenSharedFolder", path, name).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder openSharedFolder(String path) {
		return new MAPIFolder(Dispatch.call(this, "OpenSharedFolder", path).toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object openSharedItem(String path) {
		return Dispatch.call(this, "OpenSharedItem", path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param context an input-parameter of type Variant
	 * @param provider an input-parameter of type Variant
	 * @return the result is of type SharingItem
	 */
	public SharingItem createSharingItem(Variant context, Variant provider) {
		return new SharingItem(Dispatch.call(this, "CreateSharingItem", context, provider).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param context an input-parameter of type Variant
	 * @return the result is of type SharingItem
	 */
	public SharingItem createSharingItem(Variant context) {
		return new SharingItem(Dispatch.call(this, "CreateSharingItem", context).toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getExchangeMailboxServerName() {
		return Dispatch.get(this, "ExchangeMailboxServerName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getExchangeMailboxServerVersion() {
		return Dispatch.get(this, "ExchangeMailboxServerVersion").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param firstEntryID an input-parameter of type String
	 * @param secondEntryID an input-parameter of type String
	 * @return the result is of type boolean
	 */
	public boolean compareEntryIDs(String firstEntryID, String secondEntryID) {
		return Dispatch.call(this, "CompareEntryIDs", firstEntryID, secondEntryID).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAutoDiscoverXml() {
		return Dispatch.get(this, "AutoDiscoverXml").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getAutoDiscoverConnectionMode() {
		return Dispatch.get(this, "AutoDiscoverConnectionMode").changeType(Variant.VariantInt).getInt();
	}

}
