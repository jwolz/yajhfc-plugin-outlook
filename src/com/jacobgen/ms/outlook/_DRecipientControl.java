/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _DRecipientControl extends Dispatch {

	public static final String componentName = "Outlook._DRecipientControl";

	public _DRecipientControl() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _DRecipientControl(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _DRecipientControl(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type byte
	 */
	public byte getEnabled() {
		return Dispatch.get(this, "Enabled").getByte();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param enabled an input-parameter of type byte
	 */
	public void setEnabled(byte enabled) {
		Dispatch.put(this, "Enabled", enabled);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getBackColor() {
		return Dispatch.get(this, "BackColor").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param backColor an input-parameter of type int
	 */
	public void setBackColor(int backColor) {
		Dispatch.put(this, "BackColor", new Variant(backColor));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getForeColor() {
		return Dispatch.get(this, "ForeColor").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param foreColor an input-parameter of type int
	 */
	public void setForeColor(int foreColor) {
		Dispatch.put(this, "ForeColor", new Variant(foreColor));
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
	 * @return the result is of type Object
	 */
	public Object getFont() {
		return Dispatch.get(this, "Font");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param font an input-parameter of type Object
	 */
	public void setFont(Object font) {
		Dispatch.put(this, "Font", font);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSpecialEffect() {
		return Dispatch.get(this, "SpecialEffect").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param specialEffect an input-parameter of type int
	 */
	public void setSpecialEffect(int specialEffect) {
		Dispatch.put(this, "SpecialEffect", new Variant(specialEffect));
	}

}
