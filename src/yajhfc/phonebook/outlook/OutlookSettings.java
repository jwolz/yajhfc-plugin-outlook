package yajhfc.phonebook.outlook;

import yajhfc.phonebook.AbstractConnectionSettings;

public class OutlookSettings extends AbstractConnectionSettings {
	public String folderID;
	public String storeID;
	
	public boolean accessEMailAndBody = false;
	public boolean accessDistributionLists = true;
	public boolean resolveDistributionLists = true;
}
