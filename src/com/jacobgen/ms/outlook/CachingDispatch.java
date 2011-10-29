package com.jacobgen.ms.outlook;

import java.util.HashMap;
import java.util.Map;

import com.jacob.com.Dispatch;

public class CachingDispatch extends Dispatch {
	public static final Map<Class<? extends Dispatch>,Map<String,Integer>> idsOfNamesPerClass = new HashMap<Class<? extends Dispatch>,Map<String,Integer>>();
	
	public int getIDOfName(String name) {
		Map<String,Integer> classMap = idsOfNamesPerClass.get(getClass());
		if (classMap == null) {
			classMap = new HashMap<String,Integer>();
			idsOfNamesPerClass.put(getClass(), classMap);
		}
		
		Integer idOfName = classMap.get(name);
		if (idOfName == null) {
			idOfName = Dispatch.getIDOfName(this, name);
			classMap.put(name, idOfName);
			System.out.println(name + " -> " + idOfName);
		}
		return idOfName;
	}
	

	public CachingDispatch() {
		super();
	}

	public CachingDispatch(Dispatch dispatchToBeDisplaced) {
		super(dispatchToBeDisplaced);
	}

	protected CachingDispatch(int pDisp) {
		super(pDisp);
	}

	public CachingDispatch(String requestedProgramId) {
		super(requestedProgramId);
	}
}
