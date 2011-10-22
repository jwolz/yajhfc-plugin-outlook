package yajhfc.phonebook.outlook;

import static yajhfc.phonebook.outlook.EntryPoint._;

public enum AddressSource {
	HOME_ONLY(_("Home only")),
	BUSINESS_ONLY(_("Business only")),
	HOME_BEFORE_BUSINESS(_("Home before Business")),
	BUSINESS_BEFORE_HOME(_("Business before Home")),
	CUSTOM(_("Custom"));
	
	private final String description;
	
	private AddressSource(String description) {
		this.description = description;
	}

	public String getDescription() {
		return description;
	}
	
	@Override
	public String toString() {
		return description;
	}
}
