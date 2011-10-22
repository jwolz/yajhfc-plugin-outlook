/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class MAPIFolder extends Dispatch {

	public static final String componentName = "Outlook.MAPIFolder";

	public MAPIFolder() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public MAPIFolder(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public MAPIFolder(String compName) {
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
	 * @return the result is of type int
	 */
	public int getDefaultItemType() {
		return Dispatch.get(this, "DefaultItemType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDefaultMessageClass() {
		return Dispatch.get(this, "DefaultMessageClass").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDescription() {
		return Dispatch.get(this, "Description").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param description an input-parameter of type String
	 */
	public void setDescription(String description) {
		Dispatch.put(this, "Description", description);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEntryID() {
		return Dispatch.get(this, "EntryID").toString();
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
	 * @return the result is of type _Items
	 */
	public _Items getItems() {
		return new _Items(Dispatch.get(this, "Items").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getName() {
		return Dispatch.get(this, "Name").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 */
	public void setName(String name) {
		Dispatch.put(this, "Name", name);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getStoreID() {
		return Dispatch.get(this, "StoreID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getUnReadItemCount() {
		return Dispatch.get(this, "UnReadItemCount").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param destinationFolder an input-parameter of type MAPIFolder
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder copyTo(MAPIFolder destinationFolder) {
		return new MAPIFolder(Dispatch.call(this, "CopyTo", destinationFolder).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void delete() {
		Dispatch.call(this, "Delete");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void display() {
		Dispatch.call(this, "Display");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param displayMode an input-parameter of type Variant
	 * @return the result is of type _Explorer
	 */
	public _Explorer getExplorer(Variant displayMode) {
		return new _Explorer(Dispatch.call(this, "GetExplorer", displayMode).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Explorer
	 */
	public _Explorer getExplorer() {
		return new _Explorer(Dispatch.call(this, "GetExplorer").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param destinationFolder an input-parameter of type MAPIFolder
	 */
	public void moveTo(MAPIFolder destinationFolder) {
		Dispatch.call(this, "MoveTo", destinationFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getUserPermissions() {
		return Dispatch.get(this, "UserPermissions");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getWebViewOn() {
		return Dispatch.get(this, "WebViewOn").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param webViewOn an input-parameter of type boolean
	 */
	public void setWebViewOn(boolean webViewOn) {
		Dispatch.put(this, "WebViewOn", new Variant(webViewOn));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getWebViewURL() {
		return Dispatch.get(this, "WebViewURL").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param webViewURL an input-parameter of type String
	 */
	public void setWebViewURL(String webViewURL) {
		Dispatch.put(this, "WebViewURL", webViewURL);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getWebViewAllowNavigation() {
		return Dispatch.get(this, "WebViewAllowNavigation").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param webViewAllowNavigation an input-parameter of type boolean
	 */
	public void setWebViewAllowNavigation(boolean webViewAllowNavigation) {
		Dispatch.put(this, "WebViewAllowNavigation", new Variant(webViewAllowNavigation));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void addToPFFavorites() {
		Dispatch.call(this, "AddToPFFavorites");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAddressBookName() {
		return Dispatch.get(this, "AddressBookName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param addressBookName an input-parameter of type String
	 */
	public void setAddressBookName(String addressBookName) {
		Dispatch.put(this, "AddressBookName", addressBookName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowAsOutlookAB() {
		return Dispatch.get(this, "ShowAsOutlookAB").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showAsOutlookAB an input-parameter of type boolean
	 */
	public void setShowAsOutlookAB(boolean showAsOutlookAB) {
		Dispatch.put(this, "ShowAsOutlookAB", new Variant(showAsOutlookAB));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFolderPath() {
		return Dispatch.get(this, "FolderPath").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fNoUI an input-parameter of type Variant
	 * @param name an input-parameter of type Variant
	 */
	public void addToFavorites(Variant fNoUI, Variant name) {
		Dispatch.call(this, "AddToFavorites", fNoUI, name);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fNoUI an input-parameter of type Variant
	 */
	public void addToFavorites(Variant fNoUI) {
		Dispatch.call(this, "AddToFavorites", fNoUI);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void addToFavorites() {
		Dispatch.call(this, "AddToFavorites");
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getInAppFolderSyncObject() {
		return Dispatch.get(this, "InAppFolderSyncObject").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param inAppFolderSyncObject an input-parameter of type boolean
	 */
	public void setInAppFolderSyncObject(boolean inAppFolderSyncObject) {
		Dispatch.put(this, "InAppFolderSyncObject", new Variant(inAppFolderSyncObject));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type View
	 */
	public View getCurrentView() {
		return new View(Dispatch.get(this, "CurrentView").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getCustomViewsOnly() {
		return Dispatch.get(this, "CustomViewsOnly").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param customViewsOnly an input-parameter of type boolean
	 */
	public void setCustomViewsOnly(boolean customViewsOnly) {
		Dispatch.put(this, "CustomViewsOnly", new Variant(customViewsOnly));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Views
	 */
	public _Views getViews() {
		return new _Views(Dispatch.get(this, "Views").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getMAPIOBJECT() {
		return Dispatch.get(this, "MAPIOBJECT");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFullFolderPath() {
		return Dispatch.get(this, "FullFolderPath").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsSharePointFolder() {
		return Dispatch.get(this, "IsSharePointFolder").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getShowItemCount() {
		return Dispatch.get(this, "ShowItemCount").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showItemCount an input-parameter of type int
	 */
	public void setShowItemCount(int showItemCount) {
		Dispatch.put(this, "ShowItemCount", new Variant(showItemCount));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Store
	 */
	public Store getStore() {
		return new Store(Dispatch.get(this, "Store").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param storageIdentifier an input-parameter of type String
	 * @param storageIdentifierType an input-parameter of type int
	 * @return the result is of type _StorageItem
	 */
	public _StorageItem getStorage(String storageIdentifier, int storageIdentifierType) {
		return new _StorageItem(Dispatch.call(this, "GetStorage", storageIdentifier, new Variant(storageIdentifierType)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type Variant
	 * @param tableContents an input-parameter of type Variant
	 * @return the result is of type Table
	 */
	public Table getTable(Variant filter, Variant tableContents) {
		return new Table(Dispatch.call(this, "GetTable", filter, tableContents).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type Variant
	 * @return the result is of type Table
	 */
	public Table getTable(Variant filter) {
		return new Table(Dispatch.call(this, "GetTable", filter).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Table
	 */
	public Table getTable() {
		return new Table(Dispatch.call(this, "GetTable").toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type PropertyAccessor
	 */
	public PropertyAccessor getPropertyAccessor() {
		return new PropertyAccessor(Dispatch.get(this, "PropertyAccessor").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type CalendarSharing
	 */
	public CalendarSharing getCalendarExporter() {
		return new CalendarSharing(Dispatch.call(this, "GetCalendarExporter").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type UserDefinedProperties
	 */
	public UserDefinedProperties getUserDefinedProperties() {
		return new UserDefinedProperties(Dispatch.get(this, "UserDefinedProperties").toDispatch());
	}

}
