package yajhfc.phonebook.outlook;

import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.SimplePhoneBookEntry;

import com.jacobgen.ms.outlook.Recipient;

public class OlRecipientPhoneBookEntry extends SimplePhoneBookEntry {
	protected final OutlookPhonebook parent;

	protected OlRecipientPhoneBookEntry(OutlookPhonebook parent,
			Recipient rec) {
		super();
		this.parent = parent;
		readRecipient(rec);
	}

	protected void readRecipient(Recipient rec) {
		setFieldUndirty(PBEntryField.Name, rec.getName());
		
		// !Important: Also update logic in OutlookDistList.createPbeForContactAndRecipient when you change this!
		String address = rec.getAddress();
		int atPos = address.indexOf('@');
		if (atPos >= 0 && atPos < address.length() - 1) {
			String domain = address.substring(atPos+1);
			if (domain.startsWith("+")) {
				// If the domain starts with a + (as in +49 12345), assume it is a fax number
				setFieldUndirty(PBEntryField.FaxNumber, domain);
			} else {
				// Assume it is a mail address
				setFieldUndirty(PBEntryField.EMailAddress, address);
			}
		} else { // Something unknown..
			setFieldUndirty(PBEntryField.Comment, address);
		}
		
		setDirty(false);
	}

	
	@Override
	public PhoneBook getParent() {
		return parent;
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
