/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _CalendarView extends Dispatch {

	public static final String componentName = "Outlook._CalendarView";

	public _CalendarView() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _CalendarView(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _CalendarView(String compName) {
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
	 */
	public void apply() {
		Dispatch.call(this, "Apply");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param saveOption an input-parameter of type int
	 * @return the result is of type View
	 */
	public View copy(String name, int saveOption) {
		return new View(Dispatch.call(this, "Copy", name, new Variant(saveOption)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @return the result is of type View
	 */
	public View copy(String name) {
		return new View(Dispatch.call(this, "Copy", name).toDispatch());
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param name an input-parameter of type String
//	 * @param saveOption an input-parameter of type int
//	 * @return the result is of type View
//	 */
//	public View copy(String name, int saveOption) {
//		View result_of_Copy = new View(Dispatch.call(this, "Copy", name, new Variant(saveOption)).toDispatch());
//
//
//		return result_of_Copy;
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void delete() {
		Dispatch.call(this, "Delete");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void reset() {
		Dispatch.call(this, "Reset");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void save() {
		Dispatch.call(this, "Save");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLanguage() {
		return Dispatch.get(this, "Language").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param language an input-parameter of type String
	 */
	public void setLanguage(String language) {
		Dispatch.put(this, "Language", language);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getLockUserChanges() {
		return Dispatch.get(this, "LockUserChanges").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param lockUserChanges an input-parameter of type boolean
	 */
	public void setLockUserChanges(boolean lockUserChanges) {
		Dispatch.put(this, "LockUserChanges", new Variant(lockUserChanges));
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
	public int getSaveOption() {
		return Dispatch.get(this, "SaveOption").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getStandard() {
		return Dispatch.get(this, "Standard").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getViewType() {
		return Dispatch.get(this, "ViewType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getXML() {
		return Dispatch.get(this, "XML").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param xML an input-parameter of type String
	 */
	public void setXML(String xML) {
		Dispatch.put(this, "XML", xML);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param date an input-parameter of type java.util.Date
	 */
	public void goToDate(java.util.Date date) {
		Dispatch.call(this, "GoToDate", new Variant(date));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFilter() {
		return Dispatch.get(this, "Filter").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 */
	public void setFilter(String filter) {
		Dispatch.put(this, "Filter", filter);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getStartField() {
		return Dispatch.get(this, "StartField").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startField an input-parameter of type String
	 */
	public void setStartField(String startField) {
		Dispatch.put(this, "StartField", startField);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEndField() {
		return Dispatch.get(this, "EndField").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param endField an input-parameter of type String
	 */
	public void setEndField(String endField) {
		Dispatch.put(this, "EndField", endField);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getCalendarViewMode() {
		return Dispatch.get(this, "CalendarViewMode").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param calendarViewMode an input-parameter of type int
	 */
	public void setCalendarViewMode(int calendarViewMode) {
		Dispatch.put(this, "CalendarViewMode", new Variant(calendarViewMode));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDayWeekTimeScale() {
		return Dispatch.get(this, "DayWeekTimeScale").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dayWeekTimeScale an input-parameter of type int
	 */
	public void setDayWeekTimeScale(int dayWeekTimeScale) {
		Dispatch.put(this, "DayWeekTimeScale", new Variant(dayWeekTimeScale));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getMonthShowEndTime() {
		return Dispatch.get(this, "MonthShowEndTime").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param monthShowEndTime an input-parameter of type boolean
	 */
	public void setMonthShowEndTime(boolean monthShowEndTime) {
		Dispatch.put(this, "MonthShowEndTime", new Variant(monthShowEndTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getBoldDatesWithItems() {
		return Dispatch.get(this, "BoldDatesWithItems").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param boldDatesWithItems an input-parameter of type boolean
	 */
	public void setBoldDatesWithItems(boolean boldDatesWithItems) {
		Dispatch.put(this, "BoldDatesWithItems", new Variant(boldDatesWithItems));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getDayWeekTimeFont() {
		return new ViewFont(Dispatch.get(this, "DayWeekTimeFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getDayWeekFont() {
		return new ViewFont(Dispatch.get(this, "DayWeekFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getMonthFont() {
		return new ViewFont(Dispatch.get(this, "MonthFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AutoFormatRules
	 */
	public AutoFormatRules getAutoFormatRules() {
		return new AutoFormatRules(Dispatch.get(this, "AutoFormatRules").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDaysInMultiDayMode() {
		return Dispatch.get(this, "DaysInMultiDayMode").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param daysInMultiDayMode an input-parameter of type int
	 */
	public void setDaysInMultiDayMode(int daysInMultiDayMode) {
		Dispatch.put(this, "DaysInMultiDayMode", new Variant(daysInMultiDayMode));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getDisplayedDates() {
		return Dispatch.get(this, "DisplayedDates");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getBoldSubjects() {
		return Dispatch.get(this, "BoldSubjects").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param boldSubjects an input-parameter of type boolean
	 */
	public void setBoldSubjects(boolean boldSubjects) {
		Dispatch.put(this, "BoldSubjects", new Variant(boldSubjects));
	}

}
