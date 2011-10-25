package yajhfc.phonebook.outlook;
/*
 * YAJHFC - Yet another Java Hylafax client
 * Copyright (C) 2011 Jonas Wolz <info@yajhfc.de>
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
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
