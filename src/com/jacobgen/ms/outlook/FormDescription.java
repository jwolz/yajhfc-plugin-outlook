/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class FormDescription extends Dispatch {

	public static final String componentName = "Outlook.FormDescription";

	public FormDescription() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public FormDescription(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public FormDescription(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Application
	 */
	public _Application getApplication() {
		return new _Application(Dispatch.get(this, "Application").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, "Class").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, "Session").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getParent() {
		return Dispatch.get(this, "Parent");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCategory() {
		return Dispatch.get(this, "Category").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param category an input-parameter of type String
	 */
	public void setCategory(String category) {
		Dispatch.put(this, "Category", category);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCategorySub() {
		return Dispatch.get(this, "CategorySub").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param categorySub an input-parameter of type String
	 */
	public void setCategorySub(String categorySub) {
		Dispatch.put(this, "CategorySub", categorySub);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getComment() {
		return Dispatch.get(this, "Comment").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param comment an input-parameter of type String
	 */
	public void setComment(String comment) {
		Dispatch.put(this, "Comment", comment);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getContactName() {
		return Dispatch.get(this, "ContactName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param contactName an input-parameter of type String
	 */
	public void setContactName(String contactName) {
		Dispatch.put(this, "ContactName", contactName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDisplayName() {
		return Dispatch.get(this, "DisplayName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param displayName an input-parameter of type String
	 */
	public void setDisplayName(String displayName) {
		Dispatch.put(this, "DisplayName", displayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getHidden() {
		return Dispatch.get(this, "Hidden").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param hidden an input-parameter of type boolean
	 */
	public void setHidden(boolean hidden) {
		Dispatch.put(this, "Hidden", new Variant(hidden));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getIcon() {
		return Dispatch.get(this, "Icon").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param icon an input-parameter of type String
	 */
	public void setIcon(String icon) {
		Dispatch.put(this, "Icon", icon);
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
	 * @return the result is of type String
	 */
	public String getMessageClass() {
		return Dispatch.get(this, "MessageClass").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMiniIcon() {
		return Dispatch.get(this, "MiniIcon").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param miniIcon an input-parameter of type String
	 */
	public void setMiniIcon(String miniIcon) {
		Dispatch.put(this, "MiniIcon", miniIcon);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getName() {
		return Dispatch.get(this, "Name").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 */
	public void setName(String name) {
		Dispatch.put(this, "Name", name);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNumber() {
		return Dispatch.get(this, "Number").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param number an input-parameter of type String
	 */
	public void setNumber(String number) {
		Dispatch.put(this, "Number", number);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getOneOff() {
		return Dispatch.get(this, "OneOff").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param oneOff an input-parameter of type boolean
	 */
	public void setOneOff(boolean oneOff) {
		Dispatch.put(this, "OneOff", new Variant(oneOff));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPassword() {
		return Dispatch.get(this, "Password").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param password an input-parameter of type String
	 */
	public void setPassword(String password) {
		Dispatch.put(this, "Password", password);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getScriptText() {
		return Dispatch.get(this, "ScriptText").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTemplate() {
		return Dispatch.get(this, "Template").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param template an input-parameter of type String
	 */
	public void setTemplate(String template) {
		Dispatch.put(this, "Template", template);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getUseWordMail() {
		return Dispatch.get(this, "UseWordMail").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param useWordMail an input-parameter of type boolean
	 */
	public void setUseWordMail(boolean useWordMail) {
		Dispatch.put(this, "UseWordMail", new Variant(useWordMail));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getVersion() {
		return Dispatch.get(this, "Version").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param version an input-parameter of type String
	 */
	public void setVersion(String version) {
		Dispatch.put(this, "Version", version);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param registry an input-parameter of type int
	 * @param folder an input-parameter of type Variant
	 */
	public void publishForm(int registry, Variant folder) {
		Dispatch.call(this, "PublishForm", new Variant(registry), folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param registry an input-parameter of type int
	 */
	public void publishForm(int registry) {
		Dispatch.call(this, "PublishForm", new Variant(registry));
	}


}
