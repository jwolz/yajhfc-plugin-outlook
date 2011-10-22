import com.jacobgen.ms.outlook.MAPIFolder;
import com.jacobgen.ms.outlook.OlDefaultFolders;
import com.jacobgen.ms.outlook._Application;
import com.jacobgen.ms.outlook._ContactItem;
import com.jacobgen.ms.outlook._Folders;
import com.jacobgen.ms.outlook._Items;
import com.jacobgen.ms.outlook._NameSpace;


public class WrapperTest {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		_Application app = new _Application("Outlook.Application");
		_NameSpace ns = app.getNamespace("MAPI");
		
		listFolders("", ns.getFolders());
		
		MAPIFolder mf = ns.getDefaultFolder(OlDefaultFolders.olFolderContacts);
		_Items it = mf.getItems();
		
		for(int i = 1; i <= it.getCount(); i++) {
		   _ContactItem cd = new _ContactItem(it.item(i).getDispatch());
			
		   System.out.println(cd.getFullName());
		}
		
		mf = ns.getFolderFromID("000000009EF519DA190C71448D73A48DB29830BC62820000");
		System.out.println(mf.getName());
		
	}

	
	private static void listFolders(String indent, _Folders folders) {
		//System.out.println(indent + "sub-items: " + folders.getCount());
		for (int i=1; i<=folders.getCount(); i++) {
			MAPIFolder mf = folders.item(i);
			System.out.println(indent + mf.getName() + ": DefaultItemType=" + mf.getDefaultItemType() + "; Path=" + mf.getFolderPath() + "; ID=" + mf.getEntryID());
			listFolders(indent + " ", mf.getFolders());
		}
	}
}
