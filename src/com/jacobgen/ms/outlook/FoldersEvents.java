/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class FoldersEvents extends Dispatch {

	public static final String componentName = "Outlook.FoldersEvents";

	public FoldersEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public FoldersEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public FoldersEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folder an input-parameter of type MAPIFolder
	 */
	public void folderAdd(MAPIFolder folder) {
		Dispatch.call(this, "FolderAdd", folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folder an input-parameter of type MAPIFolder
	 */
	public void folderChange(MAPIFolder folder) {
		Dispatch.call(this, "FolderChange", folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void folderRemove() {
		Dispatch.call(this, "FolderRemove");
	}

}
