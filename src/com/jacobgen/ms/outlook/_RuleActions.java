/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _RuleActions extends Dispatch {

	public static final String componentName = "Outlook._RuleActions";

	public _RuleActions() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _RuleActions(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _RuleActions(String compName) {
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
	 * @param index an input-parameter of type int
	 * @return the result is of type _RuleAction
	 */
	public _RuleAction item(int index) {
		return new _RuleAction(Dispatch.call(this, "Item", new Variant(index)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MoveOrCopyRuleAction
	 */
	public MoveOrCopyRuleAction getCopyToFolder() {
		return new MoveOrCopyRuleAction(Dispatch.get(this, "CopyToFolder").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getDeletePermanently() {
		return new RuleAction(Dispatch.get(this, "DeletePermanently").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getDelete() {
		return new RuleAction(Dispatch.get(this, "Delete").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getDesktopAlert() {
		return new RuleAction(Dispatch.get(this, "DesktopAlert").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getNotifyDelivery() {
		return new RuleAction(Dispatch.get(this, "NotifyDelivery").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getNotifyRead() {
		return new RuleAction(Dispatch.get(this, "NotifyRead").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getStop() {
		return new RuleAction(Dispatch.get(this, "Stop").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MoveOrCopyRuleAction
	 */
	public MoveOrCopyRuleAction getMoveToFolder() {
		return new MoveOrCopyRuleAction(Dispatch.get(this, "MoveToFolder").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SendRuleAction
	 */
	public SendRuleAction getCC() {
		return new SendRuleAction(Dispatch.get(this, "CC").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SendRuleAction
	 */
	public SendRuleAction getForward() {
		return new SendRuleAction(Dispatch.get(this, "Forward").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SendRuleAction
	 */
	public SendRuleAction getForwardAsAttachment() {
		return new SendRuleAction(Dispatch.get(this, "ForwardAsAttachment").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SendRuleAction
	 */
	public SendRuleAction getRedirect() {
		return new SendRuleAction(Dispatch.get(this, "Redirect").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AssignToCategoryRuleAction
	 */
	public AssignToCategoryRuleAction getAssignToCategory() {
		return new AssignToCategoryRuleAction(Dispatch.get(this, "AssignToCategory").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type PlaySoundRuleAction
	 */
	public PlaySoundRuleAction getPlaySound() {
		return new PlaySoundRuleAction(Dispatch.get(this, "PlaySound").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MarkAsTaskRuleAction
	 */
	public MarkAsTaskRuleAction getMarkAsTask() {
		return new MarkAsTaskRuleAction(Dispatch.get(this, "MarkAsTask").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type NewItemAlertRuleAction
	 */
	public NewItemAlertRuleAction getNewItemAlert() {
		return new NewItemAlertRuleAction(Dispatch.get(this, "NewItemAlert").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleAction
	 */
	public RuleAction getClearCategories() {
		return new RuleAction(Dispatch.get(this, "ClearCategories").toDispatch());
	}

}
