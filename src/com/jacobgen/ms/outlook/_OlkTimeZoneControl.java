/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _OlkTimeZoneControl extends Dispatch {

	public static final String componentName = "Outlook._OlkTimeZoneControl";

	public _OlkTimeZoneControl() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _OlkTimeZoneControl(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _OlkTimeZoneControl(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getAppointmentTimeField() {
		return Dispatch.get(this, "AppointmentTimeField").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param appointmentTimeField an input-parameter of type int
	 */
	public void setAppointmentTimeField(int appointmentTimeField) {
		Dispatch.put(this, "AppointmentTimeField", new Variant(appointmentTimeField));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getBorderStyle() {
		return Dispatch.get(this, "BorderStyle").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param borderStyle an input-parameter of type int
	 */
	public void setBorderStyle(int borderStyle) {
		Dispatch.put(this, "BorderStyle", new Variant(borderStyle));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getEnabled() {
		return Dispatch.get(this, "Enabled").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param enabled an input-parameter of type boolean
	 */
	public void setEnabled(boolean enabled) {
		Dispatch.put(this, "Enabled", new Variant(enabled));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getLocked() {
		return Dispatch.get(this, "Locked").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param locked an input-parameter of type boolean
	 */
	public void setLocked(boolean locked) {
		Dispatch.put(this, "Locked", new Variant(locked));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type Picture
//	 */
//	public Picture getMouseIcon() {
//		return new Picture(Dispatch.get(this, "MouseIcon").toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param mouseIcon an input-parameter of type Picture
//	 */
//	public void setMouseIcon(Picture mouseIcon) {
//		Dispatch.put(this, "MouseIcon", mouseIcon);
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMousePointer() {
		return Dispatch.get(this, "MousePointer").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mousePointer an input-parameter of type int
	 */
	public void setMousePointer(int mousePointer) {
		Dispatch.put(this, "MousePointer", new Variant(mousePointer));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSelectedTimeZoneIndex() {
		return Dispatch.get(this, "SelectedTimeZoneIndex").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param selectedTimeZoneIndex an input-parameter of type int
	 */
	public void setSelectedTimeZoneIndex(int selectedTimeZoneIndex) {
		Dispatch.put(this, "SelectedTimeZoneIndex", new Variant(selectedTimeZoneIndex));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getValue() {
		return Dispatch.get(this, "Value");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param value an input-parameter of type Variant
	 */
	public void setValue(Variant value) {
		Dispatch.put(this, "Value", value);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void dropDown() {
		Dispatch.call(this, "DropDown");
	}

}
