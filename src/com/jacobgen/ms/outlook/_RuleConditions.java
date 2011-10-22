/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _RuleConditions extends Dispatch {

	public static final String componentName = "Outlook._RuleConditions";

	public _RuleConditions() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _RuleConditions(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _RuleConditions(String compName) {
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
	 * @return the result is of type _RuleCondition
	 */
	public _RuleCondition item(int index) {
		return new _RuleCondition(Dispatch.call(this, "Item", new Variant(index)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getCC() {
		return new RuleCondition(Dispatch.get(this, "CC").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getHasAttachment() {
		return new RuleCondition(Dispatch.get(this, "HasAttachment").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ImportanceRuleCondition
	 */
	public ImportanceRuleCondition getImportance() {
		return new ImportanceRuleCondition(Dispatch.get(this, "Importance").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getMeetingInviteOrUpdate() {
		return new RuleCondition(Dispatch.get(this, "MeetingInviteOrUpdate").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getNotTo() {
		return new RuleCondition(Dispatch.get(this, "NotTo").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getOnlyToMe() {
		return new RuleCondition(Dispatch.get(this, "OnlyToMe").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getToMe() {
		return new RuleCondition(Dispatch.get(this, "ToMe").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getToOrCc() {
		return new RuleCondition(Dispatch.get(this, "ToOrCc").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AccountRuleCondition
	 */
	public AccountRuleCondition getAccount() {
		return new AccountRuleCondition(Dispatch.get(this, "Account").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TextRuleCondition
	 */
	public TextRuleCondition getBody() {
		return new TextRuleCondition(Dispatch.get(this, "Body").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TextRuleCondition
	 */
	public TextRuleCondition getBodyOrSubject() {
		return new TextRuleCondition(Dispatch.get(this, "BodyOrSubject").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type CategoryRuleCondition
	 */
	public CategoryRuleCondition getCategory() {
		return new CategoryRuleCondition(Dispatch.get(this, "Category").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type FormNameRuleCondition
	 */
	public FormNameRuleCondition getFormName() {
		return new FormNameRuleCondition(Dispatch.get(this, "FormName").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ToOrFromRuleCondition
	 */
	public ToOrFromRuleCondition getFrom() {
		return new ToOrFromRuleCondition(Dispatch.get(this, "From").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TextRuleCondition
	 */
	public TextRuleCondition getMessageHeader() {
		return new TextRuleCondition(Dispatch.get(this, "MessageHeader").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressRuleCondition
	 */
	public AddressRuleCondition getRecipientAddress() {
		return new AddressRuleCondition(Dispatch.get(this, "RecipientAddress").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressRuleCondition
	 */
	public AddressRuleCondition getSenderAddress() {
		return new AddressRuleCondition(Dispatch.get(this, "SenderAddress").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type SenderInAddressListRuleCondition
	 */
	public SenderInAddressListRuleCondition getSenderInAddressList() {
		return new SenderInAddressListRuleCondition(Dispatch.get(this, "SenderInAddressList").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TextRuleCondition
	 */
	public TextRuleCondition getSubject() {
		return new TextRuleCondition(Dispatch.get(this, "Subject").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ToOrFromRuleCondition
	 */
	public ToOrFromRuleCondition getSentTo() {
		return new ToOrFromRuleCondition(Dispatch.get(this, "SentTo").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getOnLocalMachine() {
		return new RuleCondition(Dispatch.get(this, "OnLocalMachine").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getOnOtherMachine() {
		return new RuleCondition(Dispatch.get(this, "OnOtherMachine").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getAnyCategory() {
		return new RuleCondition(Dispatch.get(this, "AnyCategory").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RuleCondition
	 */
	public RuleCondition getFromAnyRSSFeed() {
		return new RuleCondition(Dispatch.get(this, "FromAnyRSSFeed").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type FromRssFeedRuleCondition
	 */
	public FromRssFeedRuleCondition getFromRssFeed() {
		return new FromRssFeedRuleCondition(Dispatch.get(this, "FromRssFeed").toDispatch());
	}

}
