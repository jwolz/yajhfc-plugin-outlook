/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _PropertyAccessor extends Dispatch {

	public static final String componentName = "Outlook._PropertyAccessor";

	public _PropertyAccessor() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _PropertyAccessor(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _PropertyAccessor(String compName) {
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
	 * @param schemaName an input-parameter of type String
	 * @return the result is of type Variant
	 */
	public Variant getProperty(String schemaName) {
		return Dispatch.call(this, "GetProperty", schemaName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param schemaName an input-parameter of type String
	 * @param value an input-parameter of type Variant
	 */
	public void setProperty(String schemaName, Variant value) {
		Dispatch.call(this, "SetProperty", schemaName, value);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param schemaNames an input-parameter of type Variant
	 * @return the result is of type Variant
	 */
	public Variant getProperties(Variant schemaNames) {
		return Dispatch.call(this, "GetProperties", schemaNames);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param schemaNames an input-parameter of type Variant
	 * @param values an input-parameter of type Variant
	 * @return the result is of type Variant
	 */
	public Variant setProperties(Variant schemaNames, Variant values) {
		return Dispatch.call(this, "SetProperties", schemaNames, values);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param value an input-parameter of type java.util.Date
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date uTCToLocalTime(java.util.Date value) {
		return Dispatch.call(this, "UTCToLocalTime", new Variant(value)).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param value an input-parameter of type java.util.Date
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date localTimeToUTC(java.util.Date value) {
		return Dispatch.call(this, "LocalTimeToUTC", new Variant(value)).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param value an input-parameter of type String
	 * @return the result is of type Variant
	 */
	public Variant stringToBinary(String value) {
		return Dispatch.call(this, "StringToBinary", value);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param value an input-parameter of type Variant
	 * @return the result is of type String
	 */
	public String binaryToString(Variant value) {
		return Dispatch.call(this, "BinaryToString", value).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param schemaName an input-parameter of type String
	 */
	public void deleteProperty(String schemaName) {
		Dispatch.call(this, "DeleteProperty", schemaName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param schemaNames an input-parameter of type Variant
	 * @return the result is of type Variant
	 */
	public Variant deleteProperties(Variant schemaNames) {
		return Dispatch.call(this, "DeleteProperties", schemaNames);
	}

}
