/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _TableView extends Dispatch {

	public static final String componentName = "Outlook._TableView";

	public _TableView() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _TableView(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _TableView(String compName) {
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
	 * @return the result is of type ViewFields
	 */
	public ViewFields getViewFields() {
		return new ViewFields(Dispatch.get(this, "ViewFields").toDispatch());
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
	 * @return the result is of type OrderFields
	 */
	public OrderFields getSortFields() {
		return new OrderFields(Dispatch.get(this, "SortFields").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMaxLinesInMultiLineView() {
		return Dispatch.get(this, "MaxLinesInMultiLineView").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param maxLinesInMultiLineView an input-parameter of type int
	 */
	public void setMaxLinesInMultiLineView(int maxLinesInMultiLineView) {
		Dispatch.put(this, "MaxLinesInMultiLineView", new Variant(maxLinesInMultiLineView));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutomaticGrouping() {
		return Dispatch.get(this, "AutomaticGrouping").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param automaticGrouping an input-parameter of type boolean
	 */
	public void setAutomaticGrouping(boolean automaticGrouping) {
		Dispatch.put(this, "AutomaticGrouping", new Variant(automaticGrouping));
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
	public boolean getAutomaticColumnSizing() {
		return Dispatch.get(this, "AutomaticColumnSizing").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param automaticColumnSizing an input-parameter of type boolean
	 */
	public void setAutomaticColumnSizing(boolean automaticColumnSizing) {
		Dispatch.put(this, "AutomaticColumnSizing", new Variant(automaticColumnSizing));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMultiLine() {
		return Dispatch.get(this, "MultiLine").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param multiLine an input-parameter of type int
	 */
	public void setMultiLine(int multiLine) {
		Dispatch.put(this, "MultiLine", new Variant(multiLine));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMultiLineWidth() {
		return Dispatch.get(this, "MultiLineWidth").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param multiLineWidth an input-parameter of type int
	 */
	public void setMultiLineWidth(int multiLineWidth) {
		Dispatch.put(this, "MultiLineWidth", new Variant(multiLineWidth));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAllowInCellEditing() {
		return Dispatch.get(this, "AllowInCellEditing").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param allowInCellEditing an input-parameter of type boolean
	 */
	public void setAllowInCellEditing(boolean allowInCellEditing) {
		Dispatch.put(this, "AllowInCellEditing", new Variant(allowInCellEditing));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowNewItemRow() {
		return Dispatch.get(this, "ShowNewItemRow").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showNewItemRow an input-parameter of type boolean
	 */
	public void setShowNewItemRow(boolean showNewItemRow) {
		Dispatch.put(this, "ShowNewItemRow", new Variant(showNewItemRow));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getGridLineStyle() {
		return Dispatch.get(this, "GridLineStyle").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param gridLineStyle an input-parameter of type int
	 */
	public void setGridLineStyle(int gridLineStyle) {
		Dispatch.put(this, "GridLineStyle", new Variant(gridLineStyle));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowItemsInGroups() {
		return Dispatch.get(this, "ShowItemsInGroups").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showItemsInGroups an input-parameter of type boolean
	 */
	public void setShowItemsInGroups(boolean showItemsInGroups) {
		Dispatch.put(this, "ShowItemsInGroups", new Variant(showItemsInGroups));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowReadingPane() {
		return Dispatch.get(this, "ShowReadingPane").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showReadingPane an input-parameter of type boolean
	 */
	public void setShowReadingPane(boolean showReadingPane) {
		Dispatch.put(this, "ShowReadingPane", new Variant(showReadingPane));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getHideReadingPaneHeaderInfo() {
		return Dispatch.get(this, "HideReadingPaneHeaderInfo").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param hideReadingPaneHeaderInfo an input-parameter of type boolean
	 */
	public void setHideReadingPaneHeaderInfo(boolean hideReadingPaneHeaderInfo) {
		Dispatch.put(this, "HideReadingPaneHeaderInfo", new Variant(hideReadingPaneHeaderInfo));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowUnreadAndFlaggedMessages() {
		return Dispatch.get(this, "ShowUnreadAndFlaggedMessages").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showUnreadAndFlaggedMessages an input-parameter of type boolean
	 */
	public void setShowUnreadAndFlaggedMessages(boolean showUnreadAndFlaggedMessages) {
		Dispatch.put(this, "ShowUnreadAndFlaggedMessages", new Variant(showUnreadAndFlaggedMessages));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getRowFont() {
		return new ViewFont(Dispatch.get(this, "RowFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getColumnFont() {
		return new ViewFont(Dispatch.get(this, "ColumnFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ViewFont
	 */
	public ViewFont getAutoPreviewFont() {
		return new ViewFont(Dispatch.get(this, "AutoPreviewFont").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getAutoPreview() {
		return Dispatch.get(this, "AutoPreview").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param autoPreview an input-parameter of type int
	 */
	public void setAutoPreview(int autoPreview) {
		Dispatch.put(this, "AutoPreview", new Variant(autoPreview));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AutoFormatRules
	 */
	public AutoFormatRules getAutoFormatRules() {
		return new AutoFormatRules(Dispatch.get(this, "AutoFormatRules").toDispatch());
	}

}
