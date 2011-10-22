/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class RecurrencePattern extends Dispatch {

	public static final String componentName = "Outlook.RecurrencePattern";

	public RecurrencePattern() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public RecurrencePattern(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public RecurrencePattern(String compName) {
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
	public int getDayOfMonth() {
		return Dispatch.get(this, "DayOfMonth").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dayOfMonth an input-parameter of type int
	 */
	public void setDayOfMonth(int dayOfMonth) {
		Dispatch.put(this, "DayOfMonth", new Variant(dayOfMonth));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDayOfWeekMask() {
		return Dispatch.get(this, "DayOfWeekMask").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dayOfWeekMask an input-parameter of type int
	 */
	public void setDayOfWeekMask(int dayOfWeekMask) {
		Dispatch.put(this, "DayOfWeekMask", new Variant(dayOfWeekMask));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDuration() {
		return Dispatch.get(this, "Duration").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param duration an input-parameter of type int
	 */
	public void setDuration(int duration) {
		Dispatch.put(this, "Duration", new Variant(duration));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getEndTime() {
		return Dispatch.get(this, "EndTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param endTime an input-parameter of type java.util.Date
	 */
	public void setEndTime(java.util.Date endTime) {
		Dispatch.put(this, "EndTime", new Variant(endTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Exceptions
	 */
	public Exceptions getExceptions() {
		return new Exceptions(Dispatch.get(this, "Exceptions").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getInstance() {
		return Dispatch.get(this, "Instance").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param instance an input-parameter of type int
	 */
	public void setInstance(int instance) {
		Dispatch.put(this, "Instance", new Variant(instance));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getInterval() {
		return Dispatch.get(this, "Interval").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param interval an input-parameter of type int
	 */
	public void setInterval(int interval) {
		Dispatch.put(this, "Interval", new Variant(interval));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMonthOfYear() {
		return Dispatch.get(this, "MonthOfYear").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param monthOfYear an input-parameter of type int
	 */
	public void setMonthOfYear(int monthOfYear) {
		Dispatch.put(this, "MonthOfYear", new Variant(monthOfYear));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getNoEndDate() {
		return Dispatch.get(this, "NoEndDate").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param noEndDate an input-parameter of type boolean
	 */
	public void setNoEndDate(boolean noEndDate) {
		Dispatch.put(this, "NoEndDate", new Variant(noEndDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getOccurrences() {
		return Dispatch.get(this, "Occurrences").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param occurrences an input-parameter of type int
	 */
	public void setOccurrences(int occurrences) {
		Dispatch.put(this, "Occurrences", new Variant(occurrences));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getPatternEndDate() {
		return Dispatch.get(this, "PatternEndDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param patternEndDate an input-parameter of type java.util.Date
	 */
	public void setPatternEndDate(java.util.Date patternEndDate) {
		Dispatch.put(this, "PatternEndDate", new Variant(patternEndDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getPatternStartDate() {
		return Dispatch.get(this, "PatternStartDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param patternStartDate an input-parameter of type java.util.Date
	 */
	public void setPatternStartDate(java.util.Date patternStartDate) {
		Dispatch.put(this, "PatternStartDate", new Variant(patternStartDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getRecurrenceType() {
		return Dispatch.get(this, "RecurrenceType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param recurrenceType an input-parameter of type int
	 */
	public void setRecurrenceType(int recurrenceType) {
		Dispatch.put(this, "RecurrenceType", new Variant(recurrenceType));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getRegenerate() {
		return Dispatch.get(this, "Regenerate").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param regenerate an input-parameter of type boolean
	 */
	public void setRegenerate(boolean regenerate) {
		Dispatch.put(this, "Regenerate", new Variant(regenerate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getStartTime() {
		return Dispatch.get(this, "StartTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startTime an input-parameter of type java.util.Date
	 */
	public void setStartTime(java.util.Date startTime) {
		Dispatch.put(this, "StartTime", new Variant(startTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startDate an input-parameter of type java.util.Date
	 * @return the result is of type AppointmentItem
	 */
	public AppointmentItem getOccurrence(java.util.Date startDate) {
		return new AppointmentItem(Dispatch.call(this, "GetOccurrence", new Variant(startDate)).toDispatch());
	}

}
