package yajhfc.phonebook.outlook;
import java.util.Map;

import com.jacob.com.Dispatch;
import com.jacobgen.ms.outlook._ContactItem;

import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.SimplePhoneBookEntry;

public class OutlookPhoneBookEntry extends SimplePhoneBookEntry {
	private static final String nullString = new String("");
	
	protected final OutlookPhonebook parent;
	protected String suffix;
	protected final Map<PBEntryField,String> propertyMap;
	protected _ContactItem contactItem;
	
	protected OutlookPhoneBookEntry(OutlookPhonebook parent,
			_ContactItem contactItem, Map<PBEntryField,String> propertyMap) {
		super();
		this.parent = parent;
		this.propertyMap = propertyMap;
		this.contactItem = contactItem;
		entryData.setAllFieldsTo(nullString);
	}
	
	public String getSuffix() {
		return suffix;
	}
	
	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}

//	protected void readContact(_ContactItem contact, Map<PBEntryField,String> propertyMap) {
//		//System.out.println(contact.getFullName() + ": " + contact.getEntryID());
//		for (PBEntryField field : PBEntryField.values()) {
//			String olProp = propertyMap.get(field);
//			if (olProp != null) {
//				setFieldUndirty(field, Dispatch.get(contact, contact.getIDOfName(olProp)).toString());
//			} else {
//				setFieldUndirty(field, "");
//			}
//		}
//		setDirty(false);
//	}

	/**
	 * Gets the given field from Outlook if it has not been already loaded
	 * @param field
	 * @return
	 */
	public String getFieldOnDemand(PBEntryField field) {
		String rv = super.getField(field);
		if (rv == nullString) {
			rv = loadField(field);
		}
		return rv;
	}

	/**
	 * Loads the given field from Outlook
	 * @param field
	 * @return
	 */
	protected String loadField(PBEntryField field) {
		String rv;
		String olProp = propertyMap.get(field);
		if (olProp != null) {
			synchronized (OutlookPhonebook.outlookLock) {
				rv = Dispatch.get(contactItem, contactItem.getIDOfName(olProp)).toString();
			}
		} else {
			rv = "";
		}
		setFieldUndirty(field, rv);
		return rv;
	}
	
	public boolean hasAddress() {
		for (PBEntryField field : parent.addressFields) {
			String s = getFieldOnDemand(field);
			if (s != null && s.length() > 0) {
				return true;
			}
		}
		return false;
	}
	
	/**
	 * Load all fields from Outlook
	 */
	public void loadFully() {
		for (PBEntryField field : propertyMap.keySet()) {
			getFieldOnDemand(field);
		}
		contactItem = null; // Give up unneeded reference
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
