import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBookEntry;
import yajhfc.phonebook.outlook.OutlookPhonebook;


public class PBTest {
	public static void main(String[] args) throws Exception {
		String desc = "outlook:folderID=000000009EF519DA190C71448D73A48DB29830BC62820000";
		OutlookPhonebook opb = new OutlookPhonebook(null);
		desc = opb.browseForPhoneBook(false);
		if (desc == null)
			return;
		
		opb.open(desc);
		
		for (PhoneBookEntry pbe : opb.getEntries()) {
			System.out.print(pbe); System.out.print(": ");
			for (PBEntryField field : PBEntryField.values()) {
				System.out.print(field + "=" + pbe.getField(field) + "; ");
			}
			System.out.println();
		}
	}
}
