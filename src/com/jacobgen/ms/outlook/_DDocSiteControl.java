/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _DDocSiteControl extends Dispatch {

	public static final String componentName = "Outlook._DDocSiteControl";

	public _DDocSiteControl() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _DDocSiteControl(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _DDocSiteControl(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type byte
	 */
	public byte getReadOnly() {
		return Dispatch.get(this, "ReadOnly").getByte();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param readOnly an input-parameter of type byte
	 */
	public void setReadOnly(byte readOnly) {
		Dispatch.put(this, "ReadOnly", readOnly);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type byte
	 */
	public byte getSuppressAttachments() {
		return Dispatch.get(this, "SuppressAttachments").getByte();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param suppressAttachments an input-parameter of type byte
	 */
	public void setSuppressAttachments(byte suppressAttachments) {
		Dispatch.put(this, "SuppressAttachments", suppressAttachments);
	}

}
