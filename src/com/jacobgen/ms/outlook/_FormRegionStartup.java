/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _FormRegionStartup extends Dispatch {

	public static final String componentName = "Outlook._FormRegionStartup";

	public _FormRegionStartup() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _FormRegionStartup(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _FormRegionStartup(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param formRegionName an input-parameter of type String
	 * @param item an input-parameter of type Object
	 * @param lCID an input-parameter of type int
	 * @param formRegionMode an input-parameter of type int
	 * @param formRegionSize an input-parameter of type int
	 * @return the result is of type Variant
	 */
	public Variant getFormRegionStorage(String formRegionName, Object item, int lCID, int formRegionMode, int formRegionSize) {
		return Dispatch.call(this, "GetFormRegionStorage", formRegionName, item, new Variant(lCID), new Variant(formRegionMode), new Variant(formRegionSize));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param formRegion an input-parameter of type FormRegion
	 */
	public void beforeFormRegionShow(FormRegion formRegion) {
		Dispatch.call(this, "BeforeFormRegionShow", formRegion);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param formRegionName an input-parameter of type String
	 * @param lCID an input-parameter of type int
	 * @return the result is of type Variant
	 */
	public Variant getFormRegionManifest(String formRegionName, int lCID) {
		return Dispatch.call(this, "GetFormRegionManifest", formRegionName, new Variant(lCID));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param formRegionName an input-parameter of type String
	 * @param lCID an input-parameter of type int
	 * @param icon an input-parameter of type int
	 * @return the result is of type Variant
	 */
	public Variant getFormRegionIcon(String formRegionName, int lCID, int icon) {
		return Dispatch.call(this, "GetFormRegionIcon", formRegionName, new Variant(lCID), new Variant(icon));
	}

}
