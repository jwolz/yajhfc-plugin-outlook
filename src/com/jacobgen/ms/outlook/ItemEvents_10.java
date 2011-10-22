/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ItemEvents_10 extends Dispatch {

	public static final String componentName = "Outlook.ItemEvents_10";

	public ItemEvents_10() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ItemEvents_10(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ItemEvents_10(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int open(boolean cancel) {
		return Dispatch.call(this, "Open", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int open(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_Open = Dispatch.call(this, "Open", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_Open;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param action an input-parameter of type Object
	 * @param response an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int customAction(Object action, Object response, boolean cancel) {
		return Dispatch.call(this, "CustomAction", action, response, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param action an input-parameter of type Object
	 * @param response an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int customAction(Object action, Object response, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_CustomAction = Dispatch.call(this, "CustomAction", action, response, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_CustomAction;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @return the result is of type int
	 */
	public int customPropertyChange(String name) {
		return Dispatch.call(this, "CustomPropertyChange", name).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param forward an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int forward(Object forward, boolean cancel) {
		return Dispatch.call(this, "Forward", forward, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param forward an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int forward(Object forward, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_Forward = Dispatch.call(this, "Forward", forward, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_Forward;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int close(boolean cancel) {
		return Dispatch.call(this, "Close", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int close(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_Close = Dispatch.call(this, "Close", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_Close;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @return the result is of type int
	 */
	public int propertyChange(String name) {
		return Dispatch.call(this, "PropertyChange", name).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int read() {
		return Dispatch.call(this, "Read").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param response an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int reply(Object response, boolean cancel) {
		return Dispatch.call(this, "Reply", response, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param response an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int reply(Object response, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_Reply = Dispatch.call(this, "Reply", response, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_Reply;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param response an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int replyAll(Object response, boolean cancel) {
		return Dispatch.call(this, "ReplyAll", response, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param response an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int replyAll(Object response, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_ReplyAll = Dispatch.call(this, "ReplyAll", response, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_ReplyAll;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int send(boolean cancel) {
		return Dispatch.call(this, "Send", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int send(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_Send = Dispatch.call(this, "Send", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_Send;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int write(boolean cancel) {
		return Dispatch.call(this, "Write", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int write(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_Write = Dispatch.call(this, "Write", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_Write;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeCheckNames(boolean cancel) {
		return Dispatch.call(this, "BeforeCheckNames", new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeCheckNames(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeCheckNames = Dispatch.call(this, "BeforeCheckNames", vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeCheckNames;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @return the result is of type int
	 */
	public int attachmentAdd(Attachment attachment) {
		return Dispatch.call(this, "AttachmentAdd", attachment).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @return the result is of type int
	 */
	public int attachmentRead(Attachment attachment) {
		return Dispatch.call(this, "AttachmentRead", attachment).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int beforeAttachmentSave(Attachment attachment, boolean cancel) {
		return Dispatch.call(this, "BeforeAttachmentSave", attachment, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int beforeAttachmentSave(Attachment attachment, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_BeforeAttachmentSave = Dispatch.call(this, "BeforeAttachmentSave", attachment, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_BeforeAttachmentSave;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeDelete(Object item, boolean cancel) {
		Dispatch.call(this, "BeforeDelete", item, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeDelete(Object item, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeDelete", item, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 */
	public void attachmentRemove(Attachment attachment) {
		Dispatch.call(this, "AttachmentRemove", attachment);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeAttachmentAdd(Attachment attachment, boolean cancel) {
		Dispatch.call(this, "BeforeAttachmentAdd", attachment, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeAttachmentAdd(Attachment attachment, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeAttachmentAdd", attachment, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeAttachmentPreview(Attachment attachment, boolean cancel) {
		Dispatch.call(this, "BeforeAttachmentPreview", attachment, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeAttachmentPreview(Attachment attachment, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeAttachmentPreview", attachment, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeAttachmentRead(Attachment attachment, boolean cancel) {
		Dispatch.call(this, "BeforeAttachmentRead", attachment, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeAttachmentRead(Attachment attachment, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeAttachmentRead", attachment, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeAttachmentWriteToTempFile(Attachment attachment, boolean cancel) {
		Dispatch.call(this, "BeforeAttachmentWriteToTempFile", attachment, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param attachment an input-parameter of type Attachment
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeAttachmentWriteToTempFile(Attachment attachment, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeAttachmentWriteToTempFile", attachment, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void unload() {
		Dispatch.call(this, "Unload");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeAutoSave(boolean cancel) {
		Dispatch.call(this, "BeforeAutoSave", new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeAutoSave(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeAutoSave", vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

}
