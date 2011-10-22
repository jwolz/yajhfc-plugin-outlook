/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Application extends Dispatch {

	public static final String componentName = "Outlook._Application";

	public _Application() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Application(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Application(String compName) {
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

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type Assistant
//	 */
//	public Assistant getAssistant() {
//		return new Assistant(Dispatch.get(this, "Assistant").toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getName() {
		return Dispatch.get(this, "Name").toString();
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
	 * @return the result is of type _Explorer
	 */
	public _Explorer activeExplorer() {
		return new _Explorer(Dispatch.call(this, "ActiveExplorer").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Inspector
	 */
	public _Inspector activeInspector() {
		return new _Inspector(Dispatch.call(this, "ActiveInspector").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param itemType an input-parameter of type int
	 * @return the result is of type Object
	 */
	public Object createItem(int itemType) {
		return Dispatch.call(this, "CreateItem", new Variant(itemType));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param templatePath an input-parameter of type String
	 * @param inFolder an input-parameter of type Variant
	 * @return the result is of type Object
	 */
	public Object createItemFromTemplate(String templatePath, Variant inFolder) {
		return Dispatch.call(this, "CreateItemFromTemplate", templatePath, inFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param templatePath an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object createItemFromTemplate(String templatePath) {
		return Dispatch.call(this, "CreateItemFromTemplate", templatePath);
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param templatePath an input-parameter of type String
//	 * @param inFolder an input-parameter of type Variant
//	 * @return the result is of type Object
//	 */
//	public Object createItemFromTemplate(String templatePath, Variant inFolder) {
//		Object result_of_CreateItemFromTemplate = Dispatch.call(this, "CreateItemFromTemplate", templatePath, inFolder);
//
//
//		return result_of_CreateItemFromTemplate;
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param objectName an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object createObject(String objectName) {
		return Dispatch.call(this, "CreateObject", objectName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type String
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getNamespace(String type) {
		return new _NameSpace(Dispatch.call(this, "GetNamespace", type).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void quit() {
		Dispatch.call(this, "Quit");
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type COMAddIns
//	 */
//	public COMAddIns getCOMAddIns() {
//		return new COMAddIns(Dispatch.get(this, "COMAddIns").toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Explorers
	 */
	public _Explorers getExplorers() {
		return new _Explorers(Dispatch.get(this, "Explorers").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Inspectors
	 */
	public _Inspectors getInspectors() {
		return new _Inspectors(Dispatch.get(this, "Inspectors").toDispatch());
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type LanguageSettings
//	 */
//	public LanguageSettings getLanguageSettings() {
//		return new LanguageSettings(Dispatch.get(this, "LanguageSettings").toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getProductCode() {
		return Dispatch.get(this, "ProductCode").toString();
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type AnswerWizard
//	 */
//	public AnswerWizard getAnswerWizard() {
//		return new AnswerWizard(Dispatch.get(this, "AnswerWizard").toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type MsoFeatureInstall
//	 */
//	public MsoFeatureInstall getFeatureInstall() {
//		return new MsoFeatureInstall(Dispatch.get(this, "FeatureInstall").toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param featureInstall an input-parameter of type MsoFeatureInstall
//	 */
//	public void setFeatureInstall(MsoFeatureInstall featureInstall) {
//		Dispatch.put(this, "FeatureInstall", featureInstall);
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object activeWindow() {
		return Dispatch.call(this, "ActiveWindow");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filePath an input-parameter of type String
	 * @param destFolderPath an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object copyFile(String filePath, String destFolderPath) {
		return Dispatch.call(this, "CopyFile", filePath, destFolderPath);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @param filter an input-parameter of type Variant
	 * @param searchSubFolders an input-parameter of type Variant
	 * @param tag an input-parameter of type Variant
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope, Variant filter, Variant searchSubFolders, Variant tag) {
		return new Search(Dispatch.call(this, "AdvancedSearch", scope, filter, searchSubFolders, tag).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @param filter an input-parameter of type Variant
	 * @param searchSubFolders an input-parameter of type Variant
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope, Variant filter, Variant searchSubFolders) {
		return new Search(Dispatch.call(this, "AdvancedSearch", scope, filter, searchSubFolders).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @param filter an input-parameter of type Variant
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope, Variant filter) {
		return new Search(Dispatch.call(this, "AdvancedSearch", scope, filter).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope) {
		return new Search(Dispatch.call(this, "AdvancedSearch", scope).toDispatch());
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param scope an input-parameter of type String
//	 * @param filter an input-parameter of type Variant
//	 * @param searchSubFolders an input-parameter of type Variant
//	 * @param tag an input-parameter of type Variant
//	 * @return the result is of type Search
//	 */
//	public Search advancedSearch(String scope, Variant filter, Variant searchSubFolders, Variant tag) {
//		Search result_of_AdvancedSearch = new Search(Dispatch.call(this, "AdvancedSearch", scope, filter, searchSubFolders, tag).toDispatch());
//
//
//		return result_of_AdvancedSearch;
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param lookInFolders an input-parameter of type String
	 * @return the result is of type boolean
	 */
	public boolean isSearchSynchronous(String lookInFolders) {
		return Dispatch.call(this, "IsSearchSynchronous", lookInFolders).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pvar an input-parameter of type Variant
	 */
	public void getNewNickNames(Variant pvar) {
		Dispatch.call(this, "GetNewNickNames", pvar);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Reminders
	 */
	public _Reminders getReminders() {
		return new _Reminders(Dispatch.get(this, "Reminders").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDefaultProfileName() {
		return Dispatch.get(this, "DefaultProfileName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsTrusted() {
		return Dispatch.get(this, "IsTrusted").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param referenceType an input-parameter of type int
	 * @return the result is of type Object
	 */
	public Object getObjectReference(Object item, int referenceType) {
		return Dispatch.call(this, "GetObjectReference", item, new Variant(referenceType));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type IAssistance
//	 */
//	public IAssistance getAssistance() {
//		return new IAssistance(Dispatch.get(this, "Assistance").toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TimeZones
	 */
	public TimeZones getTimeZones() {
		return new TimeZones(Dispatch.get(this, "TimeZones").toDispatch());
	}

}
