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
import static yajhfc.phonebook.outlook.EntryPoint._;
import info.clearthought.layout.TableLayout;

import java.awt.Component;
import java.awt.Dialog;
import java.awt.Font;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.List;

import javax.swing.Action;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.JTree;
import javax.swing.UIManager;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.tree.DefaultTreeCellRenderer;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import yajhfc.Utils;
import yajhfc.util.CancelAction;
import yajhfc.util.ExcDialogAbstractAction;

import com.jacobgen.ms.outlook.MAPIFolder;
import com.jacobgen.ms.outlook.OlDefaultFolders;
import com.jacobgen.ms.outlook.OlItemType;
import com.jacobgen.ms.outlook._Application;
import com.jacobgen.ms.outlook._Folders;
import com.jacobgen.ms.outlook._NameSpace;

public class ConnectionDialog extends JDialog {
	private static final int border = 6;
	
	JTree folderTree;
	Action actSetFolder, actOK;
	JTextField textSelectedFolder;
	JCheckBox checkAccessEMailAndBody, checkAccessDistList, checkResolveDistList, checkOnlyLoadFaxContacts;
	
	OutlookTreeNode selectedNode = null, rootNode;
	
	_Application app;
	_NameSpace ns;
	
	boolean clickedOK;
	
	public ConnectionDialog(Dialog owner) {
		super(owner, _("Outlook phonebook"), true);
		initialize();
	}

	public ConnectionDialog(Frame owner) {
		super(owner, _("Outlook phonebook"), true);
		initialize();
	}

	private void initialize() {
		double[][] dLay = {
			{border, TableLayout.FILL, border, 0.5, border},
			{border, TableLayout.PREFERRED, border, TableLayout.PREFERRED, TableLayout.PREFERRED, border, TableLayout.PREFERRED, border, TableLayout.PREFERRED, border, TableLayout.PREFERRED, border, TableLayout.PREFERRED, border, TableLayout.PREFERRED, TableLayout.FILL, border, TableLayout.PREFERRED, border, TableLayout.PREFERRED,  border}
		};
		
		try {
			app = new _Application("Outlook.Application");
			ns = app.getNamespace("MAPI");
		} catch (UnsatisfiedLinkError ule) {
			throw new RuntimeException("Cannot initialize COM connection: " + ule.getMessage(), ule);
		}
		
		rootNode = listOutlookFolders(ns, false);
		folderTree = new JTree(rootNode);
		folderTree.getSelectionModel().setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);
		folderTree.addTreeSelectionListener(new TreeSelectionListener() {
			@Override
			public void valueChanged(TreeSelectionEvent e) {
				OutlookTreeNode selNode = (OutlookTreeNode)e.getPath().getLastPathComponent();
				actSetFolder.setEnabled(selNode.isSelectable());
			}
		});
		folderTree.setCellRenderer(new DefaultTreeCellRenderer() {
			@Override
			public Component getTreeCellRendererComponent(JTree tree,
					Object value, boolean sel, boolean expanded, boolean leaf,
					int row, boolean hasFocus) {
				OutlookTreeNode node = (OutlookTreeNode)value;
				Component res = super.getTreeCellRendererComponent(tree, value, sel, expanded, leaf,
						row, hasFocus);
				if (!hasFocus && !node.isSelectable()) {
					res.setForeground(UIManager.getColor("textInactiveText"));
				}
				if (value == selectedNode) {
					res.setFont(folderTree.getFont().deriveFont(Font.BOLD));
				} else {
					res.setFont(folderTree.getFont());
				}
				return res;
			}
		});
		folderTree.addMouseListener(new MouseAdapter() {
            @Override
             public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                    actSetFolder.actionPerformed(null);
                }
             } 
         });
		
		textSelectedFolder = new JTextField();
		textSelectedFolder.setEditable(false);
		
		checkAccessEMailAndBody = new JCheckBox(_("Read email address and comment"));
		checkAccessDistList = new JCheckBox(_("Read distribution lists"));
		checkAccessDistList.addItemListener(new ItemListener() {
			@Override
			public void itemStateChanged(ItemEvent e) {
				checkResolveDistList.setEnabled(checkAccessDistList.isSelected());
			}
		});
		checkResolveDistList = new JCheckBox(_("Resolve distribution list items to full contacts"));
		checkResolveDistList.setEnabled(false);
		
		checkOnlyLoadFaxContacts = new JCheckBox(_("Only load contacts having a fax number"));
		
		actSetFolder = new ExcDialogAbstractAction(_("Choose folder")) {
			@Override
			protected void actualActionPerformed(ActionEvent e) {
				setSelectedNode((OutlookTreeNode)folderTree.getSelectionPath().getLastPathComponent());
			}
		};
		actSetFolder.setEnabled(false);
		
		actOK = new ExcDialogAbstractAction("OK") {
			@Override
			protected void actualActionPerformed(ActionEvent e) {
				clickedOK = true;
				setVisible(false);
			}
		};
		actOK.setEnabled(false);
		
		CancelAction actCancel = new CancelAction(this);
		
		JPanel contentPane = new JPanel(new TableLayout(dLay));
		contentPane.add(new JLabel(_("Please select which Outlook folder should be used as data source for this phonebook")), "1,1,3,1,c,c");
		contentPane.add(new JScrollPane(folderTree), "1,3,1,19,f,f");
		Utils.addWithLabel(contentPane, textSelectedFolder, _("Chosen folder:"), "3,4");
		contentPane.add(new JButton(actSetFolder), "3,6");
		contentPane.add(checkAccessEMailAndBody, "3,8");
		contentPane.add(checkAccessDistList, "3,10");
		contentPane.add(checkResolveDistList, "3,12");
		contentPane.add(checkOnlyLoadFaxContacts, "3,14");
		
		contentPane.add(new JButton(actOK), "3,17");
		contentPane.add(new JButton(actCancel), "3,19");
	
		setContentPane(contentPane);
		pack();
		setDefaultCloseOperation(HIDE_ON_CLOSE);
		setLocationByPlatform(true);
	}
	
	protected void setSelectedNode(OutlookTreeNode selNode) {
		if (selNode != null && selNode.isSelectable()) {
			selectedNode = selNode;
			textSelectedFolder.setText(selNode.folder.getFolderPath());
		}
		actOK.setEnabled(selectedNode != null && selectedNode.isSelectable());
		folderTree.repaint();
	}
	
	protected TreePath findTreeNodeByID(String entryID, String storeID) {
		OutlookTreeNode[] nodes = findTreeNodeByID(rootNode, 0, entryID, storeID);
		if (nodes != null)
			return new TreePath(nodes);
		else
			return null;
	}
	
	protected OutlookTreeNode[] findTreeNodeByID(OutlookTreeNode root, int depth, String entryID, String storeID) {
		if (root.folder != null && entryID.equals(root.folder.getEntryID()) && storeID.equals(root.folder.getStoreID())) {
			OutlookTreeNode[] res = new OutlookTreeNode[depth + 1];
			res[depth] = root;
			return res;
		}
		for (OutlookTreeNode child : root.children) {
			OutlookTreeNode[] res = findTreeNodeByID(child, depth+1, entryID, storeID);
			if (res != null) {
				res[depth] = root;
				return res;
			}
		}
		return null;
	}
	
	protected OutlookTreeNode listOutlookFolders(_NameSpace ns, boolean listAllFolders) {
		List<OutlookTreeNode> children = new ArrayList<ConnectionDialog.OutlookTreeNode>();
		_Folders folders = ns.getFolders();
		for (int i=1; i<=folders.getCount(); i++) {
			OutlookTreeNode child = listOutlookFolders(folders.item(i), listAllFolders);
			if (child != null)
				children.add(child);
		}
		OutlookTreeNode rv = new OutlookTreeNode(children);
		for (OutlookTreeNode node : children) {
			node.parent = rv;
		}
		return rv;

	}
	
	protected OutlookTreeNode listOutlookFolders(MAPIFolder mf, boolean listAllFolders) {
		List<OutlookTreeNode> children = new ArrayList<ConnectionDialog.OutlookTreeNode>();
		_Folders folders = mf.getFolders();
		for (int i=1; i<=folders.getCount(); i++) {
			OutlookTreeNode child = listOutlookFolders(folders.item(i), listAllFolders);
			if (child != null)
				children.add(child);
		}
		if (listAllFolders || children.size() > 0 || mf.getDefaultItemType() == OlItemType.olContactItem) {
			OutlookTreeNode rv = new OutlookTreeNode(mf, children);
			for (OutlookTreeNode node : children) {
				node.parent = rv;
			}
			return rv;
		} else {
			return null;
		}
	}
	
	protected void readFromConnectionSettings(OutlookSettings source) {
		TreePath selPath = null;
		if (source != null && source.folderID != null && source.folderID.length() > 0) {
			selPath = findTreeNodeByID(source.folderID, source.storeID);
		}
		if (selPath == null) { // Select default folder
			MAPIFolder mf = ns.getDefaultFolder(OlDefaultFolders.olFolderContacts);
			selPath = findTreeNodeByID(mf.getEntryID(), mf.getStoreID());
		}
		if (selPath != null) {
			folderTree.setSelectionPath(selPath);
			setSelectedNode((OutlookTreeNode)selPath.getLastPathComponent());
		} else {
			setSelectedNode(null);
		}
		
		checkAccessEMailAndBody.setSelected(source.accessEMailAndBody);
		checkAccessDistList.setSelected(source.accessDistributionLists);
		checkResolveDistList.setSelected(source.resolveDistributionLists);
		checkOnlyLoadFaxContacts.setSelected(source.loadOnlyFaxContacts);
	}
	
	protected void writeToConnectionSettings(OutlookSettings dest) {
		dest.folderID = selectedNode.folder.getEntryID();
		dest.storeID = selectedNode.folder.getStoreID();
		dest.accessEMailAndBody = checkAccessEMailAndBody.isSelected();
		dest.accessDistributionLists = checkAccessDistList.isSelected();
		dest.resolveDistributionLists = checkResolveDistList.isSelected();
		dest.loadOnlyFaxContacts = checkOnlyLoadFaxContacts.isSelected();
	}
	
    /**
     * Shows the dialog (initialized with the values of target)
     * and writes the user input into target if the user clicks "OK"
     * @param target
     * @return true if the user clicked "OK"
     */
    public boolean promptForNewSettings(OutlookSettings target) {
        readFromConnectionSettings(target);
        clickedOK = false;
        setVisible(true);
        if (clickedOK) {
            writeToConnectionSettings(target);
        }
        dispose();
        return clickedOK;
    }
	
	static class OutlookTreeNode implements TreeNode {
		public final List<OutlookTreeNode> children;
		public OutlookTreeNode parent;
		public final MAPIFolder folder;
		public final String caption;
		public final int itemType;
		
		@Override
		public TreeNode getChildAt(int childIndex) {
			return children.get(childIndex);
		}

		@Override
		public int getChildCount() {
			return (children == null) ? 0 : children.size();
		}

		@Override
		public TreeNode getParent() {
			return parent;
		}

		@Override
		public int getIndex(TreeNode node) {
			return children.indexOf(node);
		}

		@Override
		public boolean getAllowsChildren() {
			return (getChildCount() > 0);
		}

		@Override
		public boolean isLeaf() {
			return (getChildCount() == 0);
		}

		@Override
		public Enumeration<?> children() {
			return Collections.enumeration(children);
		}
		
		public boolean isSelectable() {
			return (itemType == OlItemType.olContactItem);
		}
		
		@Override
		public String toString() {
			return caption;
		}

		/**
		 * Construct a root node
		 */
		public OutlookTreeNode(List<OutlookTreeNode> children) {
			this.children = children;
			this.folder = null;
			this.caption = _("Outlook folders");
			this.itemType = -1;
		}
		
		public OutlookTreeNode(MAPIFolder folder, List<OutlookTreeNode> children) {
			super();
			this.folder = folder;
			this.caption = folder.getName();
			this.itemType = folder.getDefaultItemType();
			this.children = children;
		}
	}
}
