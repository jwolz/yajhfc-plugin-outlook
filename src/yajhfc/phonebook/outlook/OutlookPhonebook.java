package yajhfc.phonebook.outlook;

import static yajhfc.phonebook.outlook.EntryPoint._;

import java.awt.Dialog;
import java.util.ArrayList;
import java.util.Collections;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;

import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.PhoneBookEntry;
import yajhfc.phonebook.PhoneBookException;
import yajhfc.util.ExceptionDialog;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacobgen.ms.outlook.MAPIFolder;
import com.jacobgen.ms.outlook._Application;
import com.jacobgen.ms.outlook._ContactItem;
import com.jacobgen.ms.outlook._Items;
import com.jacobgen.ms.outlook._NameSpace;

public class OutlookPhonebook extends PhoneBook {
	
	public static String PB_Prefix = "outlook";      // The prefix of this Phonebook type's descriptor
	public static String PB_DisplayName = _("Microsoft Outlook contacts"); // A user-readable name for this Phonebook type
	public static String PB_Description = _("Phone book reading it entries from Microsoft Outlook contacts"); // A user-readable description of this Phonebook type

	protected static final PBEntryField[] addressFields = {
		PBEntryField.Street,
		PBEntryField.Location,
		PBEntryField.VoiceNumber,
		PBEntryField.FaxNumber,
	};
	
	protected final Map<PBEntryField,String> propertyMap_Home  = new EnumMap<PBEntryField,String>(PBEntryField.class);
	protected final Map<PBEntryField,String> propertyMap_Business  = new EnumMap<PBEntryField,String>(PBEntryField.class);
	protected final Map<PBEntryField,String> propertyMap_Other  = new EnumMap<PBEntryField,String>(PBEntryField.class);
	
	protected List<OutlookPhoneBookEntry> entries = new ArrayList<OutlookPhoneBookEntry>();
	protected List<PhoneBookEntry> entryView = Collections.<PhoneBookEntry>unmodifiableList(entries);
	
	protected OutlookSettings settings;
	protected boolean open;
	protected String folderName;
	
	public OutlookPhonebook(Dialog parent) {
		super(parent);
	}
	
	
	protected void loadBusinessAddressMapping(Map<PBEntryField,String> propertyMap) {
		propertyMap.clear();
		
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.Comment, "Body");
		propertyMap.put(PBEntryField.Company, "CompanyName");
		propertyMap.put(PBEntryField.Country, "BusinessAddressCountry");
		propertyMap.put(PBEntryField.Department, "Department");
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.EMailAddress, "Email1Address");
		propertyMap.put(PBEntryField.FaxNumber, "BusinessFaxNumber");
		propertyMap.put(PBEntryField.GivenName, "FirstName");
		propertyMap.put(PBEntryField.Location, "BusinessAddressCity");
		propertyMap.put(PBEntryField.Name, "LastName");
		propertyMap.put(PBEntryField.Position, "JobTitle"); 
		propertyMap.put(PBEntryField.State, "BusinessAddressState");
		propertyMap.put(PBEntryField.Street, "BusinessAddressStreet");
		propertyMap.put(PBEntryField.Title, "Title");
		propertyMap.put(PBEntryField.VoiceNumber, "BusinessTelephoneNumber");
		propertyMap.put(PBEntryField.WebSite, "BusinessHomePage");
		propertyMap.put(PBEntryField.ZIPCode, "BusinessAddressPostalCode");
	}
	
	protected void loadHomeAddressMapping(Map<PBEntryField,String> propertyMap) {
		propertyMap.clear();
		
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.Comment, "Body");
		//propertyMap.put(PBEntryField.Company, "CompanyName");
		propertyMap.put(PBEntryField.Country, "HomeAddressCountry");
		//propertyMap.put(PBEntryField.Department, "Department");
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.EMailAddress, "Email2Address");
		propertyMap.put(PBEntryField.FaxNumber, "HomeFaxNumber");
		propertyMap.put(PBEntryField.GivenName, "FirstName");
		propertyMap.put(PBEntryField.Location, "HomeAddressCity");
		propertyMap.put(PBEntryField.Name, "LastName");
		//propertyMap.put(PBEntryField.Position, "JobTitle"); 
		propertyMap.put(PBEntryField.State, "HomeAddressState");
		propertyMap.put(PBEntryField.Street, "HomeAddressStreet");
		propertyMap.put(PBEntryField.Title, "Title");
		propertyMap.put(PBEntryField.VoiceNumber, "HomeTelephoneNumber");
		propertyMap.put(PBEntryField.WebSite, "WebPage");
		propertyMap.put(PBEntryField.ZIPCode, "HomeAddressPostalCode");
	}

	protected void loadOtherAddressMapping(Map<PBEntryField,String> propertyMap) {
		propertyMap.clear();
		
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.Comment, "Body");
		//propertyMap.put(PBEntryField.Company, "CompanyName");
		propertyMap.put(PBEntryField.Country, "OtherAddressCountry");
		//propertyMap.put(PBEntryField.Department, "Department");
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.EMailAddress, "Email3Address");
		propertyMap.put(PBEntryField.FaxNumber, "OtherFaxNumber");
		propertyMap.put(PBEntryField.GivenName, "FirstName");
		propertyMap.put(PBEntryField.Location, "OtherAddressCity");
		propertyMap.put(PBEntryField.Name, "LastName");
		//propertyMap.put(PBEntryField.Position, "JobTitle"); 
		propertyMap.put(PBEntryField.State, "OtherAddressState");
		propertyMap.put(PBEntryField.Street, "OtherAddressStreet");
		propertyMap.put(PBEntryField.Title, "Title");
		propertyMap.put(PBEntryField.VoiceNumber, "OtherTelephoneNumber");
		//propertyMap.put(PBEntryField.WebSite, "OtherHomePage");
		propertyMap.put(PBEntryField.ZIPCode, "OtherAddressPostalCode");
	}
	
	@Override
	public PhoneBookEntry addNewEntry() {
		return null;
	}

	@Override
	public List<PhoneBookEntry> getEntries() {
		return entryView;
	}

	@Override
	public String browseForPhoneBook(boolean exportMode) {
		try {
			OutlookSettings newSettings = new OutlookSettings();
			if (settings != null)
				newSettings.copyFrom(settings);
			ConnectionDialog cd = new ConnectionDialog(parentDialog);
			if (cd.promptForNewSettings(newSettings)) {
				return PB_Prefix + ":" + newSettings.saveToString();
			} else {
				return null;
			}
		} catch (Exception e) {
			ExceptionDialog.showExceptionDialog(parentDialog, _("Cannot connect to Outlook"), e);
			return null;
		}
	}

	@Override
	public boolean isOpen() {
		return open;
	}

	@Override
	public String getDisplayCaption() {
		return "Outlook - " + folderName;
	}
	
	@Override
	public void resort() {
		Collections.sort(entries);
	}

	@Override
	protected void openInternal(String descriptorWithoutPrefix)
			throws PhoneBookException {
		settings = new OutlookSettings();
		settings.loadFromString(descriptorWithoutPrefix);
		
		loadBusinessAddressMapping(propertyMap_Business);
		loadHomeAddressMapping(propertyMap_Home);
		loadOtherAddressMapping(propertyMap_Other);
		
		_Application app = new _Application("Outlook.Application");
		_NameSpace ns = app.getNamespace("MAPI");
		
		MAPIFolder contacts = ns.getFolderFromID(settings.folderID, new Variant(settings.storeID));
		folderName = contacts.getName();
		_Items cl = contacts.getItems();
		entries.clear();
		for (int i=1; i<=cl.getCount(); i++) {
			_ContactItem ci = new _ContactItem(cl.item(i).getDispatch());
			int numAddresses = 0;
			boolean haveBusiness = false, haveHome = false, haveOther = false;
			if (haveAddress(ci, propertyMap_Business)) {
				haveBusiness = true;
				numAddresses++;
			}
			if (haveAddress(ci, propertyMap_Home)) {
				haveHome = true;
				numAddresses++;
			}
			if (haveAddress(ci, propertyMap_Other)) {
				haveOther = true;
				numAddresses++;
			}
			if (haveBusiness || numAddresses == 0) {
				entries.add(new OutlookPhoneBookEntry(this, ci, propertyMap_Business, (numAddresses > 1) ? _("Business") : null));	
			}
			if (haveHome) {
				entries.add(new OutlookPhoneBookEntry(this, ci, propertyMap_Home, (numAddresses > 1) ? _("Home") : null));	
			}
			if (haveOther) {
				entries.add(new OutlookPhoneBookEntry(this, ci, propertyMap_Other, (numAddresses > 1) ? _("Other") : null));	
			}
		}
		open = true;
	}
	
	protected boolean haveAddress(_ContactItem ci, Map<PBEntryField,String> propertyMap) {
		for (PBEntryField field : addressFields) {
			String olProp = propertyMap.get(field);
			if (olProp != null) {
				String s = Dispatch.get(ci, olProp).toString();
				if (s != null && s.length() > 0) {
					return true;
				}
			}
		}
		return false;
	}

	@Override
	public void close() {
		open = false;
	}

	@Override
	public boolean isReadOnly() {
		return true;
	}
	
	
}
