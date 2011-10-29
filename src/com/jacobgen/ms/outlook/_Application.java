/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Application extends CachingDispatch {

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
		return new _Application(Dispatch.get(this, getIDOfName("Application")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, getIDOfName("Class")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, getIDOfName("Session")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getParent() {
		return Dispatch.get(this, getIDOfName("Parent"));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type Assistant
//	 */
//	public Assistant getAssistant() {
//		return new Assistant(Dispatch.get(this, getIDOfName("Assistant")).toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getName() {
		return Dispatch.get(this, getIDOfName("Name")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getVersion() {
		return Dispatch.get(this, getIDOfName("Version")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Explorer
	 */
	public _Explorer activeExplorer() {
		return new _Explorer(Dispatch.call(this, getIDOfName("ActiveExplorer")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Inspector
	 */
	public _Inspector activeInspector() {
		return new _Inspector(Dispatch.call(this, getIDOfName("ActiveInspector")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param itemType an input-parameter of type int
	 * @return the result is of type Object
	 */
	public Object createItem(int itemType) {
		return Dispatch.call(this, getIDOfName("CreateItem"), new Variant(itemType));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param templatePath an input-parameter of type String
	 * @param inFolder an input-parameter of type Variant
	 * @return the result is of type Object
	 */
	public Object createItemFromTemplate(String templatePath, Variant inFolder) {
		return Dispatch.call(this, getIDOfName("CreateItemFromTemplate"), templatePath, inFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param templatePath an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object createItemFromTemplate(String templatePath) {
		return Dispatch.call(this, getIDOfName("CreateItemFromTemplate"), templatePath);
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param templatePath an input-parameter of type String
//	 * @param inFolder an input-parameter of type Variant
//	 * @return the result is of type Object
//	 */
//	public Object createItemFromTemplate(String templatePath, Variant inFolder) {
//		Object result_of_CreateItemFromTemplate = Dispatch.call(this, getIDOfName("CreateItemFromTemplate"), templatePath, inFolder);
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
		return Dispatch.call(this, getIDOfName("CreateObject"), objectName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type String
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getNamespace(String type) {
		return new _NameSpace(Dispatch.call(this, getIDOfName("GetNamespace"), type).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void quit() {
		Dispatch.call(this, getIDOfName("Quit"));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type COMAddIns
//	 */
//	public COMAddIns getCOMAddIns() {
//		return new COMAddIns(Dispatch.get(this, getIDOfName("COMAddIns")).toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Explorers
	 */
	public _Explorers getExplorers() {
		return new _Explorers(Dispatch.get(this, getIDOfName("Explorers")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Inspectors
	 */
	public _Inspectors getInspectors() {
		return new _Inspectors(Dispatch.get(this, getIDOfName("Inspectors")).toDispatch());
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type LanguageSettings
//	 */
//	public LanguageSettings getLanguageSettings() {
//		return new LanguageSettings(Dispatch.get(this, getIDOfName("LanguageSettings")).toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getProductCode() {
		return Dispatch.get(this, getIDOfName("ProductCode")).toString();
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type AnswerWizard
//	 */
//	public AnswerWizard getAnswerWizard() {
//		return new AnswerWizard(Dispatch.get(this, getIDOfName("AnswerWizard")).toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type MsoFeatureInstall
//	 */
//	public MsoFeatureInstall getFeatureInstall() {
//		return new MsoFeatureInstall(Dispatch.get(this, getIDOfName("FeatureInstall")).toDispatch());
//	}
//
//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @param featureInstall an input-parameter of type MsoFeatureInstall
//	 */
//	public void setFeatureInstall(MsoFeatureInstall featureInstall) {
//		Dispatch.put(this, getIDOfName("FeatureInstall"), featureInstall);
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object activeWindow() {
		return Dispatch.call(this, getIDOfName("ActiveWindow"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filePath an input-parameter of type String
	 * @param destFolderPath an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Object copyFile(String filePath, String destFolderPath) {
		return Dispatch.call(this, getIDOfName("CopyFile"), filePath, destFolderPath);
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
		return new Search(Dispatch.call(this, getIDOfName("AdvancedSearch"), scope, filter, searchSubFolders, tag).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @param filter an input-parameter of type Variant
	 * @param searchSubFolders an input-parameter of type Variant
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope, Variant filter, Variant searchSubFolders) {
		return new Search(Dispatch.call(this, getIDOfName("AdvancedSearch"), scope, filter, searchSubFolders).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @param filter an input-parameter of type Variant
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope, Variant filter) {
		return new Search(Dispatch.call(this, getIDOfName("AdvancedSearch"), scope, filter).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param scope an input-parameter of type String
	 * @return the result is of type Search
	 */
	public Search advancedSearch(String scope) {
		return new Search(Dispatch.call(this, getIDOfName("AdvancedSearch"), scope).toDispatch());
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
//		Search result_of_AdvancedSearch = new Search(Dispatch.call(this, getIDOfName("AdvancedSearch"), scope, filter, searchSubFolders, tag).toDispatch());
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
		return Dispatch.call(this, getIDOfName("IsSearchSynchronous"), lookInFolders).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pvar an input-parameter of type Variant
	 */
	public void getNewNickNames(Variant pvar) {
		Dispatch.call(this, getIDOfName("GetNewNickNames"), pvar);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Reminders
	 */
	public _Reminders getReminders() {
		return new _Reminders(Dispatch.get(this, getIDOfName("Reminders")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDefaultProfileName() {
		return Dispatch.get(this, getIDOfName("DefaultProfileName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsTrusted() {
		return Dispatch.get(this, getIDOfName("IsTrusted")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param referenceType an input-parameter of type int
	 * @return the result is of type Object
	 */
	public Object getObjectReference(Object item, int referenceType) {
		return Dispatch.call(this, getIDOfName("GetObjectReference"), item, new Variant(referenceType));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
//	 * @return the result is of type IAssistance
//	 */
//	public IAssistance getAssistance() {
//		return new IAssistance(Dispatch.get(this, getIDOfName("Assistance")).toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TimeZones
	 */
	public TimeZones getTimeZones() {
		return new TimeZones(Dispatch.get(this, getIDOfName("TimeZones")).toDispatch());
	}

}
