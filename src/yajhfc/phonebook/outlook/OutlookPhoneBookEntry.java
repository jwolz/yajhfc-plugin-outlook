package yajhfc.phonebook.outlook;

import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.SimplePhoneBookEntry;

import com.jacob.com.Dispatch;
import com.jacobgen.ms.outlook._ContactItem;

public class OutlookPhoneBookEntry extends SimplePhoneBookEntry {
	
	
	protected final OutlookPhonebook parent;

	protected OutlookPhoneBookEntry(OutlookPhonebook parent,
			_ContactItem contactItem) {
		super();
		this.parent = parent;
		readContact(contactItem);
	}

	protected void readContact(_ContactItem contact) {
		for (PBEntryField field : PBEntryField.values()) {
			String olProp = parent.settings.getMappingFor(field);
			if (olProp != null && !olProp.equals(OutlookSettings.noField)) {
				setFieldUndirty(field, Dispatch.get(contact, olProp).toString());
			} else {
				setFieldUndirty(field, "");
			}
		}
		setDirty(false);
	}
	
	
	@Override
	public PhoneBook getParent() {
		return parent;
	}

	@Override
	public void delete() {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void commit() {
		// TODO Auto-generated method stub
		
	}


}
