/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ExplorerEvents_10 extends Dispatch {

	public static final String componentName = "Outlook.ExplorerEvents_10";

	public ExplorerEvents_10() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ExplorerEvents_10(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ExplorerEvents_10(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int activate() {
		return Dispatch.call(this, "Activate").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int folderSwitch() {
		return Dispatch.call(this, "FolderSwitch").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param newFolder an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeFolderSwitch(Object newFolder, boolean cancel) {
		return Dispatch.call(this, "BeforeFolderSwitch", newFolder, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param newFolder an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeFolderSwitch(Object newFolder, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeFolderSwitch = Dispatch.call(this, "BeforeFolderSwitch", newFolder, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeFolderSwitch;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int viewSwitch() {
		return Dispatch.call(this, "ViewSwitch").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param newView an input-parameter of type Variant
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeViewSwitch(Variant newView, boolean cancel) {
		return Dispatch.call(this, "BeforeViewSwitch", newView, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param newView an input-parameter of type Variant
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeViewSwitch(Variant newView, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeViewSwitch = Dispatch.call(this, "BeforeViewSwitch", newView, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeViewSwitch;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int deactivate() {
		return Dispatch.call(this, "Deactivate").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int selectionChange() {
		return Dispatch.call(this, "SelectionChange").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int close() {
		return Dispatch.call(this, "Close").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeMaximize(boolean cancel) {
		return Dispatch.call(this, "BeforeMaximize", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeMaximize(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeMaximize = Dispatch.call(this, "BeforeMaximize", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeMaximize;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeMinimize(boolean cancel) {
		return Dispatch.call(this, "BeforeMinimize", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeMinimize(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeMinimize = Dispatch.call(this, "BeforeMinimize", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeMinimize;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeMove(boolean cancel) {
		return Dispatch.call(this, "BeforeMove", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeMove(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeMove = Dispatch.call(this, "BeforeMove", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeMove;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeSize(boolean cancel) {
		return Dispatch.call(this, "BeforeSize", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeSize(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeSize = Dispatch.call(this, "BeforeSize", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeSize;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeItemCopy(boolean cancel) {
		Dispatch.call(this, "BeforeItemCopy", new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeItemCopy(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeItemCopy", vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeItemCut(boolean cancel) {
		Dispatch.call(this, "BeforeItemCut", new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeItemCut(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeItemCut", vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param clipboardContent an input-parameter of type Variant
	 * @param target an input-parameter of type MAPIFolder
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeItemPaste(Variant clipboardContent, MAPIFolder target, boolean cancel) {
		Dispatch.call(this, "BeforeItemPaste", clipboardContent, target, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param clipboardContent an input-parameter of type Variant
	 * @param target an input-parameter of type MAPIFolder
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeItemPaste(Variant clipboardContent, MAPIFolder target, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeItemPaste", clipboardContent, target, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

}
