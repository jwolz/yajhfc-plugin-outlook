package yajhfc.phonebook.outlook;

import java.util.Map;

import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.SimplePhoneBookEntry;

import com.jacob.com.Dispatch;
import com.jacobgen.ms.outlook._ContactItem;

public class OutlookPhoneBookEntry extends SimplePhoneBookEntry {
	
	
	protected final OutlookPhonebook parent;
	protected final String suffix;

	protected OutlookPhoneBookEntry(OutlookPhonebook parent,
			_ContactItem contactItem, Map<PBEntryField,String> propertyMap, String suffix) {
		super();
		this.parent = parent;
		this.suffix = suffix;
		readContact(contactItem, propertyMap);
	}

	protected void readContact(_ContactItem contact, Map<PBEntryField,String> propertyMap) {
		//System.out.println(contact.getFullName() + ": " + contact.getEntryID());
		for (PBEntryField field : PBEntryField.values()) {
			String olProp = propertyMap.get(field);
			if (olProp != null) {
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
	public String toString() {
		if (suffix != null) {
			return super.toString() + " (" + suffix + ")";
		} else {
			return super.toString();
		}
	}

	@Override
	public void delete() {
		// Not modifiable
	}

	@Override
	public void commit() {
		//  Not modifiable
	}

	

}
