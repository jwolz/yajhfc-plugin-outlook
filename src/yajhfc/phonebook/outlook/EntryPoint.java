package yajhfc.phonebook.outlook;

import java.util.logging.Logger;

import yajhfc.launch.Launcher2;
import yajhfc.phonebook.PhoneBookFactory;
import yajhfc.phonebook.PhoneBookType;
import yajhfc.plugin.PluginManager;

public class EntryPoint {

	private static final Logger log = Logger.getLogger(EntryPoint.class.getName());
	
	public static String _(String key) {
		return key;
	}
	
	
//	public static File getJacobDir() {
//		// Try to determine where the JAR file is located
//		URL utilURL = Dispatch.class.getResource("Dispatch.class");
//		try {
//			while (utilURL.getProtocol().equals("jar")) {
//				String path = utilURL.getPath();
//				int idx = path.lastIndexOf('!');
//				if (idx >= 0) {
//					path = path.substring(0, idx);
//				}
//				utilURL = new URL(path);
//			}
//		} catch (MalformedURLException e) {
//			log.log(Level.WARNING, "Error determining application dir:", e);
//		}
//		if (utilURL.getProtocol().equals("file")) {
//			try {
//				URI uri = utilURL.toURI();
//				if (Utils.IS_WINDOWS && uri.getAuthority() != null) {
//					// Work around a JDK bug with UNC paths
//					uri = new URI("file", null, "////" + uri.getAuthority() + '/' + uri.getPath(), null); 
//				}
//				return (new File(uri)).getParentFile();
//			} catch (URISyntaxException e) {
//				log.log(Level.SEVERE, "JACOB directory not found, url was: " +  Dispatch.class.getResource("Dispatch.class"), e);
//				return null;
//			}
//		} else {
//			log.severe("JACOB directory not found, url was: " +  Dispatch.class.getResource("Dispatch.class"));
//			return null;
//		}
//
//	}
	
	/**
	 * Plugin initialization method.
	 * The name and signature of this method must be exactly as follows 
	 * (i.e. it must always be "public static boolean init(int)" )
	 * @param startupMode the mode YajHFC is starting up in. The possible
	 *    values are one of the STARTUP_MODE_* constants defined in yajhfc.plugin.PluginManager
	 * @return true if the initialization was successful, false otherwise.
	 */
	public static boolean init(int startupMode) {
		PhoneBookFactory.PhonebookTypes.add(new PhoneBookType(OutlookPhonebook.class));
		return true;
	}
	
    /**
     * Launches YajHFC including this plugin (for debugging purposes)
     * @param args
     */
    public static void main(String[] args) {
		PluginManager.internalPlugins.add(EntryPoint.class);
		Launcher2.main(args);
	}
}
