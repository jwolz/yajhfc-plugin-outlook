/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ApplicationEvents_11 extends Dispatch {

	public static final String componentName = "Outlook.ApplicationEvents_11";

	public ApplicationEvents_11() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ApplicationEvents_11(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ApplicationEvents_11(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int itemSend(Object item, boolean cancel) {
		return Dispatch.call(this, "ItemSend", item, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int itemSend(Object item, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_ItemSend = Dispatch.call(this, "ItemSend", item, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_ItemSend;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int newMail() {
		return Dispatch.call(this, "NewMail").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @return the result is of type int
	 */
	public int reminder(Object item) {
		return Dispatch.call(this, "Reminder", item).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pages an input-parameter of type PropertyPages
	 * @return the result is of type int
	 */
	public int optionsPagesAdd(PropertyPages pages) {
		return Dispatch.call(this, "OptionsPagesAdd", pages).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int startup() {
		return Dispatch.call(this, "Startup").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int quit() {
		return Dispatch.call(this, "Quit").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param searchObject an input-parameter of type Search
	 * @return the result is of type int
	 */
	public int advancedSearchComplete(Search searchObject) {
		return Dispatch.call(this, "AdvancedSearchComplete", searchObject).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param searchObject an input-parameter of type Search
	 * @return the result is of type int
	 */
	public int advancedSearchStopped(Search searchObject) {
		return Dispatch.call(this, "AdvancedSearchStopped", searchObject).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int mAPILogonComplete() {
		return Dispatch.call(this, "MAPILogonComplete").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param entryIDCollection an input-parameter of type String
	 */
	public void newMailEx(String entryIDCollection) {
		Dispatch.call(this, "NewMailEx", entryIDCollection);
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param commandBar an input-parameter of type CommandBar
//	 * @param attachments an input-parameter of type AttachmentSelection
//	 * @return the result is of type int
//	 */
//	public int attachmentContextMenuDisplay(CommandBar commandBar, AttachmentSelection attachments) {
//		return Dispatch.call(this, "AttachmentContextMenuDisplay", commandBar, attachments).changeType(Variant.VariantInt).getInt();
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param commandBar an input-parameter of type CommandBar
//	 * @param folder an input-parameter of type Folder
//	 */
//	public void folderContextMenuDisplay(CommandBar commandBar, Folder folder) {
//		Dispatch.call(this, "FolderContextMenuDisplay", commandBar, folder);
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param commandBar an input-parameter of type CommandBar
//	 * @param store an input-parameter of type Store
//	 */
//	public void storeContextMenuDisplay(CommandBar commandBar, Store store) {
//		Dispatch.call(this, "StoreContextMenuDisplay", commandBar, store);
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param commandBar an input-parameter of type CommandBar
//	 * @param shortcut an input-parameter of type OutlookBarShortcut
//	 */
//	public void shortcutContextMenuDisplay(CommandBar commandBar, OutlookBarShortcut shortcut) {
//		Dispatch.call(this, "ShortcutContextMenuDisplay", commandBar, shortcut);
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param commandBar an input-parameter of type CommandBar
//	 * @param view an input-parameter of type View
//	 */
//	public void viewContextMenuDisplay(CommandBar commandBar, View view) {
//		Dispatch.call(this, "ViewContextMenuDisplay", commandBar, view);
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param commandBar an input-parameter of type CommandBar
//	 * @param selection an input-parameter of type Selection
//	 */
//	public void itemContextMenuDisplay(CommandBar commandBar, Selection selection) {
//		Dispatch.call(this, "ItemContextMenuDisplay", commandBar, selection);
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param contextMenu an input-parameter of type int
	 */
	public void contextMenuClose(int contextMenu) {
		Dispatch.call(this, "ContextMenuClose", new Variant(contextMenu));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 */
	public void itemLoad(Object item) {
		Dispatch.call(this, "ItemLoad", item);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folderToShare an input-parameter of type MAPIFolder
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeFolderSharingDialog(MAPIFolder folderToShare, boolean cancel) {
		Dispatch.call(this, "BeforeFolderSharingDialog", folderToShare, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param folderToShare an input-parameter of type MAPIFolder
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeFolderSharingDialog(MAPIFolder folderToShare, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeFolderSharingDialog", folderToShare, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

}
