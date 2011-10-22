/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _OlkTextBox extends Dispatch {

	public static final String componentName = "Outlook._OlkTextBox";

	public _OlkTextBox() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _OlkTextBox(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _OlkTextBox(String compName) {
		super(compName);
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

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutoTab() {
		return Dispatch.get(this, "AutoTab").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param autoTab an input-parameter of type boolean
	 */
	public void setAutoTab(boolean autoTab) {
		Dispatch.put(this, "AutoTab", new Variant(autoTab));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutoWordSelect() {
		return Dispatch.get(this, "AutoWordSelect").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param autoWordSelect an input-parameter of type boolean
	 */
	public void setAutoWordSelect(boolean autoWordSelect) {
		Dispatch.put(this, "AutoWordSelect", new Variant(autoWordSelect));
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
	 * @return the result is of type int
	 */
	public int getDragBehavior() {
		return Dispatch.get(this, "DragBehavior").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dragBehavior an input-parameter of type int
	 */
	public void setDragBehavior(int dragBehavior) {
		Dispatch.put(this, "DragBehavior", new Variant(dragBehavior));
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
	 * @return the result is of type int
	 */
	public int getEnterFieldBehavior() {
		return Dispatch.get(this, "EnterFieldBehavior").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param enterFieldBehavior an input-parameter of type int
	 */
	public void setEnterFieldBehavior(int enterFieldBehavior) {
		Dispatch.put(this, "EnterFieldBehavior", new Variant(enterFieldBehavior));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getEnterKeyBehavior() {
		return Dispatch.get(this, "EnterKeyBehavior").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param enterKeyBehavior an input-parameter of type boolean
	 */
	public void setEnterKeyBehavior(boolean enterKeyBehavior) {
		Dispatch.put(this, "EnterKeyBehavior", new Variant(enterKeyBehavior));
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

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getHideSelection() {
		return Dispatch.get(this, "HideSelection").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param hideSelection an input-parameter of type boolean
	 */
	public void setHideSelection(boolean hideSelection) {
		Dispatch.put(this, "HideSelection", new Variant(hideSelection));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIntegralHeight() {
		return Dispatch.get(this, "IntegralHeight").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param integralHeight an input-parameter of type boolean
	 */
	public void setIntegralHeight(boolean integralHeight) {
		Dispatch.put(this, "IntegralHeight", new Variant(integralHeight));
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

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMaxLength() {
		return Dispatch.get(this, "MaxLength").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param maxLength an input-parameter of type int
	 */
	public void setMaxLength(int maxLength) {
		Dispatch.put(this, "MaxLength", new Variant(maxLength));
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
	 * @return the result is of type boolean
	 */
	public boolean getMultiLine() {
		return Dispatch.get(this, "MultiLine").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param multiLine an input-parameter of type boolean
	 */
	public void setMultiLine(boolean multiLine) {
		Dispatch.put(this, "MultiLine", new Variant(multiLine));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPasswordChar() {
		return Dispatch.get(this, "PasswordChar").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param passwordChar an input-parameter of type String
	 */
	public void setPasswordChar(String passwordChar) {
		Dispatch.put(this, "PasswordChar", passwordChar);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getScrollbars() {
		return Dispatch.get(this, "Scrollbars").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scrollbars an input-parameter of type int
	 */
	public void setScrollbars(int scrollbars) {
		Dispatch.put(this, "Scrollbars", new Variant(scrollbars));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getSelectionMargin() {
		return Dispatch.get(this, "SelectionMargin").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param selectionMargin an input-parameter of type boolean
	 */
	public void setSelectionMargin(boolean selectionMargin) {
		Dispatch.put(this, "SelectionMargin", new Variant(selectionMargin));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getTabKeyBehavior() {
		return Dispatch.get(this, "TabKeyBehavior").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param tabKeyBehavior an input-parameter of type boolean
	 */
	public void setTabKeyBehavior(boolean tabKeyBehavior) {
		Dispatch.put(this, "TabKeyBehavior", new Variant(tabKeyBehavior));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getText() {
		return Dispatch.get(this, "Text").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param text an input-parameter of type String
	 */
	public void setText(String text) {
		Dispatch.put(this, "Text", text);
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
	 * @return the result is of type int
	 */
	public int getSelStart() {
		return Dispatch.get(this, "SelStart").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param selStart an input-parameter of type int
	 */
	public void setSelStart(int selStart) {
		Dispatch.put(this, "SelStart", new Variant(selStart));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSelLength() {
		return Dispatch.get(this, "SelLength").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param selLength an input-parameter of type int
	 */
	public void setSelLength(int selLength) {
		Dispatch.put(this, "SelLength", new Variant(selLength));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSelText() {
		return Dispatch.get(this, "SelText").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void cut() {
		Dispatch.call(this, "Cut");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void copy() {
		Dispatch.call(this, "Copy");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void paste() {
		Dispatch.call(this, "Paste");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void clear() {
		Dispatch.call(this, "Clear");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getEnableRichText() {
		return Dispatch.get(this, "EnableRichText").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param enableRichText an input-parameter of type boolean
	 */
	public void setEnableRichText(boolean enableRichText) {
		Dispatch.put(this, "EnableRichText", new Variant(enableRichText));
	}

}
