/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class OlkComboBoxEvents extends Dispatch {

	public static final String componentName = "Outlook.OlkComboBoxEvents";

	public OlkComboBoxEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public OlkComboBoxEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public OlkComboBoxEvents(String compName) {
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

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void enter() {
		Dispatch.call(this, "Enter");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 */
	public void exit(boolean cancel) {
		Dispatch.call(this, "Exit", new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param keyCode an input-parameter of type int
	 * @param shift an input-parameter of type int
	 */
	public void keyDown(int keyCode, int shift) {
		Dispatch.call(this, "KeyDown", new Variant(keyCode), new Variant(shift));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param keyAscii an input-parameter of type int
	 */
	public void keyPress(int keyAscii) {
		Dispatch.call(this, "KeyPress", new Variant(keyAscii));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param keyCode an input-parameter of type int
	 * @param shift an input-parameter of type int
	 */
	public void keyUp(int keyCode, int shift) {
		Dispatch.call(this, "KeyUp", new Variant(keyCode), new Variant(shift));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void change() {
		Dispatch.call(this, "Change");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void dropButtonClick() {
		Dispatch.call(this, "DropButtonClick");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void afterUpdate() {
		Dispatch.call(this, "AfterUpdate");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeUpdate(boolean cancel) {
		Dispatch.call(this, "BeforeUpdate", new Variant(cancel));
	}

}
