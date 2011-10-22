import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Variant;


public class COMTest {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		ActiveXComponent ol = new ActiveXComponent("Outlook.Application");
		
		ActiveXComponent namespace = ol.invokeGetComponent("GetNamespace", new Variant("MAPI"));
		
		ActiveXComponent defFolder = namespace.invokeGetComponent("GetDefaultFolder", new Variant(10));
		
		ActiveXComponent items = defFolder.getPropertyAsComponent("Items");
		
		for(int i = 1; i <= items.getPropertyAsInt("Count"); i++)
		{
			ActiveXComponent item = items.invokeGetComponent("Item", new Variant(i));
			
			System.out.println(item.invoke("FullName").getString());
		}
		
		
	}

}
