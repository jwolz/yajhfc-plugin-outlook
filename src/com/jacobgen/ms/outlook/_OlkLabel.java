/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _OlkLabel extends Dispatch {

	public static final String componentName = "Outlook._OlkLabel";

	public _OlkLabel() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _OlkLabel(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _OlkLabel(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAccelerator() {
		return Dispatch.get(this, "Accelerator").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param accelerator an input-parameter of type String
	 */
	public void setAccelerator(String accelerator) {
		Dispatch.put(this, "Accelerator", accelerator);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutoSize() {
		return Dispatch.get(this, "AutoSize").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param autoSize an input-parameter of type boolean
	 */
	public void setAutoSize(boolean autoSize) {
		Dispatch.put(this, "AutoSize", new Variant(autoSize));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type OLE_COLOR
//	 */
//	public OLE_COLOR getBackColor() {
//		return new OLE_COLOR(Dispatch.get(this, "BackColor").toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param backColor an input-parameter of type OLE_COLOR
//	 */
//	public void setBackColor(OLE_COLOR backColor) {
//		Dispatch.put(this, "BackColor", backColor);
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getBackStyle() {
		return Dispatch.get(this, "BackStyle").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param backStyle an input-parameter of type int
	 */
	public void setBackStyle(int backStyle) {
		Dispatch.put(this, "BackStyle", new Variant(backStyle));
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
	 * @return the result is of type String
	 */
	public String getCaption() {
		return Dispatch.get(this, "Caption").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param caption an input-parameter of type String
	 */
	public void setCaption(String caption) {
		Dispatch.put(this, "Caption", caption);
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

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type Font
//	 */
//	public Font getFont() {
//		return new Font(Dispatch.get(this, "Font").toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type OLE_COLOR
//	 */
//	public OLE_COLOR getForeColor() {
//		return new OLE_COLOR(Dispatch.get(this, "ForeColor").toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param foreColor an input-parameter of type OLE_COLOR
//	 */
//	public void setForeColor(OLE_COLOR foreColor) {
//		Dispatch.put(this, "ForeColor", foreColor);
//	}
//
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
	public int getTextAlign() {
		return Dispatch.get(this, "TextAlign").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param textAlign an input-parameter of type int
	 */
	public void setTextAlign(int textAlign) {
		Dispatch.put(this, "TextAlign", new Variant(textAlign));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getWordWrap() {
		return Dispatch.get(this, "WordWrap").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param wordWrap an input-parameter of type boolean
	 */
	public void setWordWrap(boolean wordWrap) {
		Dispatch.put(this, "WordWrap", new Variant(wordWrap));
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
	 * @return the result is of type boolean
	 */
	public boolean getUseHeaderColor() {
		return Dispatch.get(this, "UseHeaderColor").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param useHeaderColor an input-parameter of type boolean
	 */
	public void setUseHeaderColor(boolean useHeaderColor) {
		Dispatch.put(this, "UseHeaderColor", new Variant(useHeaderColor));
	}

}
