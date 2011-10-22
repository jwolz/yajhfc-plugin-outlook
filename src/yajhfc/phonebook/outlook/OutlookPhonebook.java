package yajhfc.phonebook.outlook;

import static yajhfc.phonebook.outlook.EntryPoint._;

import java.awt.Dialog;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.PhoneBookEntry;
import yajhfc.phonebook.PhoneBookException;

import com.jacobgen.ms.outlook.MAPIFolder;
import com.jacobgen.ms.outlook._Application;
import com.jacobgen.ms.outlook._ContactItem;
import com.jacobgen.ms.outlook._Items;
import com.jacobgen.ms.outlook._NameSpace;

public class OutlookPhonebook extends PhoneBook {
	
	public static String PB_Prefix = "outlook";      // The prefix of this Phonebook type's descriptor
	public static String PB_DisplayName = _("Microsoft Outlook"); // A user-readable name for this Phonebook type
	public static String PB_Description = _("Phone book reading it entries from Microsoft Outlook contacts"); // A user-readable description of this Phonebook type

	
	protected List<OutlookPhoneBookEntry> entries = new ArrayList<OutlookPhoneBookEntry>();
	protected List<PhoneBookEntry> entryView = Collections.<PhoneBookEntry>unmodifiableList(entries);
	
	protected OutlookSettings settings;
	
	public OutlookPhonebook(Dialog parent) {
		super(parent);
		// TODO Auto-generated constructor stub
	}

	@Override
	public PhoneBookEntry addNewEntry() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<PhoneBookEntry> getEntries() {
		return entryView;
	}

	@Override
	public String browseForPhoneBook(boolean exportMode) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public boolean isOpen() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void resort() {
		// TODO Auto-generated method stub

	}

	@Override
	protected void openInternal(String descriptorWithoutPrefix)
			throws PhoneBookException {
		settings = new OutlookSettings();
		settings.loadFromString(descriptorWithoutPrefix);
		
		
		_Application app = new _Application("Outlook.Application");
		_NameSpace ns = app.getNamespace("MAPI");
		
		MAPIFolder contacts = ns.getFolderFromID(settings.folderID);
		_Items cl = contacts.getItems();
		entries.clear();
		for (int i=1; i<=cl.getCount(); i++) {
			_ContactItem ci = new _ContactItem(cl.item(i).getDispatch());
			entries.add(new OutlookPhoneBookEntry(this, ci));
		}
	}

	@Override
	public void close() {
		// TODO Auto-generated method stub

	}

}
