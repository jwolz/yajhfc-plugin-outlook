/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class OlkLabelEvents extends Dispatch {

	public static final String componentName = "Outlook.OlkLabelEvents";

	public OlkLabelEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public OlkLabelEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public OlkLabelEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void click() {
		Dispatch.call(this, "Click");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void doubleClick() {
		Dispatch.call(this, "DoubleClick");
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param button an input-parameter of type int
//	 * @param shift an input-parameter of type int
//	 * @param x an input-parameter of type OLE_XPOS_CONTAINER
//	 * @param y an input-parameter of type OLE_YPOS_CONTAINER
//	 */
//	public void mouseDown(int button, int shift, OLE_XPOS_CONTAINER x, OLE_YPOS_CONTAINER y) {
//		Dispatch.call(this, "MouseDown", new Variant(button), new Variant(shift), x, y);
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param button an input-parameter of type int
//	 * @param shift an input-parameter of type int
//	 * @param x an input-parameter of type OLE_XPOS_CONTAINER
//	 * @param y an input-parameter of type OLE_YPOS_CONTAINER
//	 */
//	public void mouseMove(int button, int shift, OLE_XPOS_CONTAINER x, OLE_YPOS_CONTAINER y) {
//		Dispatch.call(this, "MouseMove", new Variant(button), new Variant(shift), x, y);
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param button an input-parameter of type int
//	 * @param shift an input-parameter of type int
//	 * @param x an input-parameter of type OLE_XPOS_CONTAINER
//	 * @param y an input-parameter of type OLE_YPOS_CONTAINER
//	 */
//	public void mouseUp(int button, int shift, OLE_XPOS_CONTAINER x, OLE_YPOS_CONTAINER y) {
//		Dispatch.call(this, "MouseUp", new Variant(button), new Variant(shift), x, y);
//	}

}
