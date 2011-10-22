/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _CalendarSharing extends Dispatch {

	public static final String componentName = "Outlook._CalendarSharing";

	public _CalendarSharing() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _CalendarSharing(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _CalendarSharing(String compName) {
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
	 * @param path an input-parameter of type String
	 */
	public void saveAsICal(String path) {
		Dispatch.call(this, "SaveAsICal", path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailFormat an input-parameter of type int
	 * @return the result is of type MailItem
	 */
	public MailItem forwardAsICal(int mailFormat) {
		return new MailItem(Dispatch.call(this, "ForwardAsICal", new Variant(mailFormat)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getCalendarDetail() {
		return Dispatch.get(this, "CalendarDetail").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param calendarDetail an input-parameter of type int
	 */
	public void setCalendarDetail(int calendarDetail) {
		Dispatch.put(this, "CalendarDetail", new Variant(calendarDetail));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getEndDate() {
		return Dispatch.get(this, "EndDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param endDate an input-parameter of type java.util.Date
	 */
	public void setEndDate(java.util.Date endDate) {
		Dispatch.put(this, "EndDate", new Variant(endDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getFolder() {
		return new MAPIFolder(Dispatch.get(this, "Folder").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIncludeAttachments() {
		return Dispatch.get(this, "IncludeAttachments").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param includeAttachments an input-parameter of type boolean
	 */
	public void setIncludeAttachments(boolean includeAttachments) {
		Dispatch.put(this, "IncludeAttachments", new Variant(includeAttachments));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIncludePrivateDetails() {
		return Dispatch.get(this, "IncludePrivateDetails").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param includePrivateDetails an input-parameter of type boolean
	 */
	public void setIncludePrivateDetails(boolean includePrivateDetails) {
		Dispatch.put(this, "IncludePrivateDetails", new Variant(includePrivateDetails));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getRestrictToWorkingHours() {
		return Dispatch.get(this, "RestrictToWorkingHours").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param restrictToWorkingHours an input-parameter of type boolean
	 */
	public void setRestrictToWorkingHours(boolean restrictToWorkingHours) {
		Dispatch.put(this, "RestrictToWorkingHours", new Variant(restrictToWorkingHours));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getStartDate() {
		return Dispatch.get(this, "StartDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startDate an input-parameter of type java.util.Date
	 */
	public void setStartDate(java.util.Date startDate) {
		Dispatch.put(this, "StartDate", new Variant(startDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIncludeWholeCalendar() {
		return Dispatch.get(this, "IncludeWholeCalendar").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param includeWholeCalendar an input-parameter of type boolean
	 */
	public void setIncludeWholeCalendar(boolean includeWholeCalendar) {
		Dispatch.put(this, "IncludeWholeCalendar", new Variant(includeWholeCalendar));
	}

}
