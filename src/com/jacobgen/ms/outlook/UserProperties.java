/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class UserProperties extends Dispatch {

	public static final String componentName = "Outlook.UserProperties";

	public UserProperties() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public UserProperties(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public UserProperties(String compName) {
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
	 * @return the result is of type int
	 */
	public int getCount() {
		return Dispatch.get(this, "Count").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 * @return the result is of type UserProperty
	 */
	public UserProperty item(Variant index) {
		return new UserProperty(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param type an input-parameter of type int
	 * @param addToFolderFields an input-parameter of type Variant
	 * @param displayFormat an input-parameter of type Variant
	 * @return the result is of type UserProperty
	 */
	public UserProperty add(String name, int type, Variant addToFolderFields, Variant displayFormat) {
		return new UserProperty(Dispatch.call(this, "Add", name, new Variant(type), addToFolderFields, displayFormat).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param type an input-parameter of type int
	 * @param addToFolderFields an input-parameter of type Variant
	 * @return the result is of type UserProperty
	 */
	public UserProperty add(String name, int type, Variant addToFolderFields) {
		return new UserProperty(Dispatch.call(this, "Add", name, new Variant(type), addToFolderFields).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param type an input-parameter of type int
	 * @return the result is of type UserProperty
	 */
	public UserProperty add(String name, int type) {
		return new UserProperty(Dispatch.call(this, "Add", name, new Variant(type)).toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param custom an input-parameter of type Variant
	 * @return the result is of type UserProperty
	 */
	public UserProperty find(String name, Variant custom) {
		return new UserProperty(Dispatch.call(this, "Find", name, custom).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @return the result is of type UserProperty
	 */
	public UserProperty find(String name) {
		return new UserProperty(Dispatch.call(this, "Find", name).toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type int
	 */
	public void remove(int index) {
		Dispatch.call(this, "Remove", new Variant(index));
	}

}
