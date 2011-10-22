/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _TimelineView extends Dispatch {

	public static final String componentName = "Outlook._TimelineView";

	public _TimelineView() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _TimelineView(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _TimelineView(String compName) {
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
	 * @return the result is of type OrderFields
	 */
	public OrderFields getGroupByFields() {
		return new OrderFields(Dispatch.get(this, "GroupByFields").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDefaultExpandCollapseSetting() {
		return Dispatch.get(this, "DefaultExpandCollapseSetting").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param defaultExpandCollapseSetting an input-parameter of type int
	 */
	public void setDefaultExpandCollapseSetting(int defaultExpandCollapseSetting) {
		Dispatch.put(this, "DefaultExpandCollapseSetting", new Variant(defaultExpandCollapseSetting));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowWeekNumbers() {
		return Dispatch.get(this, "ShowWeekNumbers").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showWeekNumbers an input-parameter of type boolean
	 */
	public void setShowWeekNumbers(boolean showWeekNumbers) {
		Dispatch.put(this, "ShowWeekNumbers", new Variant(showWeekNumbers));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMaxLabelWidth() {
		return Dispatch.get(this, "MaxLabelWidth").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param maxLabelWidth an input-parameter of type int
	 */
	public void setMaxLabelWidth(int maxLabelWidth) {
		Dispatch.put(this, "MaxLabelWidth", new Variant(maxLabelWidth));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowLabelWhenViewingByMonth() {
		return Dispatch.get(this, "ShowLabelWhenViewingByMonth").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showLabelWhenViewingByMonth an input-parameter of type boolean
	 */
	public void setShowLabelWhenViewingByMonth(boolean showLabelWhenViewingByMonth) {
		Dispatch.put(this, "ShowLabelWhenViewingByMonth", new Variant(showLabelWhenViewingByMonth));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getUpperScaleFont() {
		return new ViewFont(Dispatch.get(this, "UpperScaleFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getLowerScaleFont() {
		return new ViewFont(Dispatch.get(this, "LowerScaleFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getItemFont() {
		return new ViewFont(Dispatch.get(this, "ItemFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getTimelineViewMode() {
		return Dispatch.get(this, "TimelineViewMode").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param timelineViewMode an input-parameter of type int
	 */
	public void setTimelineViewMode(int timelineViewMode) {
		Dispatch.put(this, "TimelineViewMode", new Variant(timelineViewMode));
	}

}
