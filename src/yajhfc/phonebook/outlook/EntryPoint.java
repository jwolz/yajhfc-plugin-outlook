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
import java.util.logging.Logger;

import yajhfc.launch.Launcher2;
import yajhfc.phonebook.PhoneBookFactory;
import yajhfc.phonebook.PhoneBookType;
import yajhfc.plugin.PluginManager;
import yajhfc.util.MsgBundle;

public class EntryPoint {	
    public static final MsgBundle msgBundle  = new MsgBundle("yajhfc.phonebook.outlook.i18n.Messages");
    
    /**
     * Returns the translation of key. If no translation is found, the
     * key is returned.
     * @param key
     * @return
     */
    public static String _(String key) {
        return msgBundle._(key, key);
    }
    
    /**
     * Returns the translation of key. If no translation is found, the
     * defaultValue is returned.
     * @param key
     * @param defaultValue
     * @return
     */
    public static String _(String key, String defaultValue) {
        return msgBundle._(key, defaultValue);
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
