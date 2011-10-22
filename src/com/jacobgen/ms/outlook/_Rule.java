/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Rule extends Dispatch {

	public static final String componentName = "Outlook._Rule";

	public _Rule() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Rule(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Rule(String compName) {
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
	 * @return the result is of type int
	 */
	public int getExecutionOrder() {
		return Dispatch.get(this, "ExecutionOrder").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param executionOrder an input-parameter of type int
	 */
	public void setExecutionOrder(int executionOrder) {
		Dispatch.put(this, "ExecutionOrder", new Variant(executionOrder));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getRuleType() {
		return Dispatch.get(this, "RuleType").changeType(Variant.VariantInt).getInt();
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
	public boolean getIsLocalRule() {
		return Dispatch.get(this, "IsLocalRule").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showProgress an input-parameter of type Variant
	 * @param folder an input-parameter of type Variant
	 * @param includeSubfolders an input-parameter of type Variant
	 * @param ruleExecuteOption an input-parameter of type Variant
	 */
	public void execute(Variant showProgress, Variant folder, Variant includeSubfolders, Variant ruleExecuteOption) {
		Dispatch.call(this, "Execute", showProgress, folder, includeSubfolders, ruleExecuteOption);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showProgress an input-parameter of type Variant
	 * @param folder an input-parameter of type Variant
	 * @param includeSubfolders an input-parameter of type Variant
	 */
	public void execute(Variant showProgress, Variant folder, Variant includeSubfolders) {
		Dispatch.call(this, "Execute", showProgress, folder, includeSubfolders);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showProgress an input-parameter of type Variant
	 * @param folder an input-parameter of type Variant
	 */
	public void execute(Variant showProgress, Variant folder) {
		Dispatch.call(this, "Execute", showProgress, folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showProgress an input-parameter of type Variant
	 */
	public void execute(Variant showProgress) {
		Dispatch.call(this, "Execute", showProgress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void execute() {
		Dispatch.call(this, "Execute");
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleActions
	 */
	public RuleActions getActions() {
		return new RuleActions(Dispatch.get(this, "Actions").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleConditions
	 */
	public RuleConditions getConditions() {
		return new RuleConditions(Dispatch.get(this, "Conditions").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleConditions
	 */
	public RuleConditions getExceptions() {
		return new RuleConditions(Dispatch.get(this, "Exceptions").toDispatch());
	}

}
