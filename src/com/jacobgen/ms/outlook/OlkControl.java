/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class OlkControl extends Dispatch {

	public static final String componentName = "Outlook.OlkControl";

	public OlkControl() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public OlkControl(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public OlkControl(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getItemProperty() {
		return Dispatch.get(this, "ItemProperty").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param itemProperty an input-parameter of type String
	 */
	public void setItemProperty(String itemProperty) {
		Dispatch.put(this, "ItemProperty", itemProperty);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getControlProperty() {
		return Dispatch.get(this, "ControlProperty").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param controlProperty an input-parameter of type String
	 */
	public void setControlProperty(String controlProperty) {
		Dispatch.put(this, "ControlProperty", controlProperty);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPossibleValues() {
		return Dispatch.get(this, "PossibleValues").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param possibleValues an input-parameter of type String
	 */
	public void setPossibleValues(String possibleValues) {
		Dispatch.put(this, "PossibleValues", possibleValues);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getFormat() {
		return Dispatch.get(this, "Format").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param format an input-parameter of type int
	 */
	public void setFormat(int format) {
		Dispatch.put(this, "Format", new Variant(format));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getEnableAutoLayout() {
		return Dispatch.get(this, "EnableAutoLayout").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param enableAutoLayout an input-parameter of type boolean
	 */
	public void setEnableAutoLayout(boolean enableAutoLayout) {
		Dispatch.put(this, "EnableAutoLayout", new Variant(enableAutoLayout));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMinimumWidth() {
		return Dispatch.get(this, "MinimumWidth").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param minimumWidth an input-parameter of type int
	 */
	public void setMinimumWidth(int minimumWidth) {
		Dispatch.put(this, "MinimumWidth", new Variant(minimumWidth));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMinimumHeight() {
		return Dispatch.get(this, "MinimumHeight").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param minimumHeight an input-parameter of type int
	 */
	public void setMinimumHeight(int minimumHeight) {
		Dispatch.put(this, "MinimumHeight", new Variant(minimumHeight));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getHorizontalLayout() {
		return Dispatch.get(this, "HorizontalLayout").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param horizontalLayout an input-parameter of type int
	 */
	public void setHorizontalLayout(int horizontalLayout) {
		Dispatch.put(this, "HorizontalLayout", new Variant(horizontalLayout));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getVerticalLayout() {
		return Dispatch.get(this, "VerticalLayout").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param verticalLayout an input-parameter of type int
	 */
	public void setVerticalLayout(int verticalLayout) {
		Dispatch.put(this, "VerticalLayout", new Variant(verticalLayout));
	}

}
