/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class InspectorEvents_10 extends Dispatch {

	public static final String componentName = "Outlook.InspectorEvents_10";

	public InspectorEvents_10() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public InspectorEvents_10(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public InspectorEvents_10(String compName) {
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
	public int deactivate() {
		return Dispatch.call(this, "Deactivate").changeType(Variant.VariantInt).getInt();
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
	 * @param activePageName an input-parameter of type String
	 */
	public void pageChange(String activePageName) {
		Dispatch.call(this, "PageChange", activePageName);
	}

}
