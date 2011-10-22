/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Explorer extends Dispatch {

	public static final String componentName = "Outlook._Explorer";

	public _Explorer() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Explorer(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Explorer(String compName) {
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
//	 * @return the result is of type CommandBars
//	 */
//	public CommandBars getCommandBars() {
//		return new CommandBars(Dispatch.get(this, "CommandBars").toDispatch());
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getCurrentFolder() {
		return new MAPIFolder(Dispatch.get(this, "CurrentFolder").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param currentFolder an input-parameter of type MAPIFolder
	 */
	public void setCurrentFolder(MAPIFolder currentFolder) {
		Dispatch.put(this, "CurrentFolder", currentFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void close() {
		Dispatch.call(this, "Close");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void display() {
		Dispatch.call(this, "Display");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCaption() {
		return Dispatch.get(this, "Caption").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getCurrentView() {
		return Dispatch.get(this, "CurrentView");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param currentView an input-parameter of type Variant
	 */
	public void setCurrentView(Variant currentView) {
		Dispatch.put(this, "CurrentView", currentView);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getHeight() {
		return Dispatch.get(this, "Height").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param height an input-parameter of type int
	 */
	public void setHeight(int height) {
		Dispatch.put(this, "Height", new Variant(height));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getLeft() {
		return Dispatch.get(this, "Left").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param left an input-parameter of type int
	 */
	public void setLeft(int left) {
		Dispatch.put(this, "Left", new Variant(left));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Panes
	 */
	public Panes getPanes() {
		return new Panes(Dispatch.get(this, "Panes").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Selection
	 */
	public Selection getSelection() {
		return new Selection(Dispatch.get(this, "Selection").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getTop() {
		return Dispatch.get(this, "Top").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param top an input-parameter of type int
	 */
	public void setTop(int top) {
		Dispatch.put(this, "Top", new Variant(top));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getWidth() {
		return Dispatch.get(this, "Width").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param width an input-parameter of type int
	 */
	public void setWidth(int width) {
		Dispatch.put(this, "Width", new Variant(width));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getWindowState() {
		return Dispatch.get(this, "WindowState").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param windowState an input-parameter of type int
	 */
	public void setWindowState(int windowState) {
		Dispatch.put(this, "WindowState", new Variant(windowState));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void activate() {
		Dispatch.call(this, "Activate");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pane an input-parameter of type int
	 * @return the result is of type boolean
	 */
	public boolean isPaneVisible(int pane) {
		return Dispatch.call(this, "IsPaneVisible", new Variant(pane)).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pane an input-parameter of type int
	 * @param visible an input-parameter of type boolean
	 */
	public void showPane(int pane, boolean visible) {
		Dispatch.call(this, "ShowPane", new Variant(pane), new Variant(visible));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getViews() {
		return Dispatch.get(this, "Views");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getHTMLDocument() {
		return Dispatch.get(this, "HTMLDocument");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folder an input-parameter of type MAPIFolder
	 */
	public void selectFolder(MAPIFolder folder) {
		Dispatch.call(this, "SelectFolder", folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folder an input-parameter of type MAPIFolder
	 */
	public void deselectFolder(MAPIFolder folder) {
		Dispatch.call(this, "DeselectFolder", folder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param folder an input-parameter of type MAPIFolder
	 * @return the result is of type boolean
	 */
	public boolean isFolderSelected(MAPIFolder folder) {
		return Dispatch.call(this, "IsFolderSelected", folder).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NavigationPane
	 */
	public _NavigationPane getNavigationPane() {
		return new _NavigationPane(Dispatch.get(this, "NavigationPane").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void clearSearch() {
		Dispatch.call(this, "ClearSearch");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param query an input-parameter of type String
	 * @param searchScope an input-parameter of type int
	 */
	public void search(String query, int searchScope) {
		Dispatch.call(this, "Search", query, new Variant(searchScope));
	}

}
