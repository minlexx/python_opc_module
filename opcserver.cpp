#include <Python.h>
#include <structmember.h>
#include "module.h"

#include "opc/opccomn.h"
#include "opc/OpcEnum.h"
#include "opc/opcda.h"


#define MY_OPC_GROUP_NAME L"PythonOPCGroup"


typedef struct opch_added_item {
	DWORD      dwAccessRights; // OPC_READABLE | OPC_WRITEABLE;
	VARTYPE    vtDataType;
	OPCHANDLE  hServerHandle;
	OPCHANDLE  hClientHandle;
    time_t     tmTimeStamp;
    WORD       wQuality;
	wchar_t   *pszItemID;
	//
	struct opch_added_item *next;
} tagopch_added_items;


typedef struct opch_IOPCServer {
	PyObject_HEAD
	// python objects - object members
	// need special caring
	PyObject          *m_szGuid;
	PyObject          *m_szProgID;
	// OPC, COM related vars
	IOPCServer        *m_srv;
	IOPCGroupStateMgt *m_groupMgt;
	IOPCItemMgt       *m_itemMgt;
	IOPCSyncIO        *m_syncIO;
	OPCHANDLE          m_hServerGroup;
	int                m_bSupportsV3;
	LCID               m_server_lang;
	// linked list of added OPC items
	struct opch_added_item *added_items;
} tagopch_IOPCServer;


static struct opch_added_item *opch_IOPCServer_find_added_item(opch_IOPCServer *self, const wchar_t *pszItemID) {
	struct opch_added_item *ret = NULL;
	struct opch_added_item *cur = self->added_items;
	if (!pszItemID) return ret;
	while (cur) {
		if (cur->pszItemID) {
			if (wcscmp(cur->pszItemID, pszItemID) == 0) {
				ret = cur;
				break;
			}
		}
		cur = cur->next;
	}
	return ret;
}


static int opch_IOPCServer_traverse(opch_IOPCServer *self, visitproc visit, void *arg) {
	Py_VISIT(self->m_szGuid);
	Py_VISIT(self->m_szProgID);
	return 0;
}


static int opch_IOPCServer_clear(opch_IOPCServer *self) {
	Py_CLEAR(self->m_szGuid);
	Py_CLEAR(self->m_szProgID);
	return 0;
}


static void opch_IOPCServer_dealloc(opch_IOPCServer *self) {
	if (self->m_syncIO) {
		self->m_syncIO->Release();
		self->m_syncIO = NULL;
	}
	if (self->m_srv && self->m_groupMgt && self->m_itemMgt && (self->m_hServerGroup > 0)) {
		self->m_itemMgt->Release();
		self->m_groupMgt->Release();
		self->m_srv->RemoveGroup(self->m_hServerGroup, TRUE);
		self->m_itemMgt = NULL;
		self->m_groupMgt = NULL;
	}
	if (self->m_srv) {
		self->m_srv->Release();
		self->m_srv = NULL;
	}
	if (self->added_items) {
		struct opch_added_item *cur_item = self->added_items;
		while (cur_item) {
			if (cur_item->pszItemID) {
				free(cur_item->pszItemID);
				cur_item->pszItemID = NULL;
			}
			struct opch_added_item *cur_tmp = cur_item;
			cur_item = cur_item->next;
			free(cur_tmp);
		}
	}
	opch_IOPCServer_clear(self);
	Py_TYPE(self)->tp_free(self);
}


static PyObject *opch_IOPCServer_new(PyTypeObject *type, PyObject *args, PyObject *kwds) {
	opch_IOPCServer *self = NULL;
	self = (opch_IOPCServer *)type->tp_alloc(type, 0);
	if (self != NULL) {
		self->m_srv = NULL;
		self->m_bSupportsV3 = FALSE;
		self->m_server_lang = MAKELANGID(LANG_NEUTRAL, SUBLANG_SYS_DEFAULT);
		self->m_groupMgt = NULL;
		self->m_itemMgt = NULL;
		self->m_hServerGroup = 0;
		// create python object members
		self->m_szGuid = PyUnicode_FromString("");
		if (self->m_szGuid == NULL) {
			Py_DECREF(self);
			return NULL;
		}
		self->m_szProgID = PyUnicode_FromString("");
		if (self->m_szProgID == NULL) {
			Py_DECREF(self);
			return NULL;
		}
	}
	return (PyObject *)self;
}

static int opch_IOPCServer_init(opch_IOPCServer *self, PyObject *args, PyObject *kwds) {
	PyObject *szGuid = NULL;
	PyObject *tmp = NULL;
	static char *kwlist[] = {"guid", NULL};
	if (!PyArg_ParseTupleAndKeywords(args, kwds, "|O", kwlist, &szGuid)) {
		return -1;
	}
	if (szGuid) {
		tmp = self->m_szGuid;
		Py_INCREF(szGuid);
		self->m_szGuid = szGuid;
		Py_XDECREF(tmp);
	}
	return 0;
}

static PyObject *opch_IOPCServer_repr(opch_IOPCServer *self) {
	if (!self) {
		return NULL;
	}
	if (self->m_srv == NULL) {
		return PyUnicode_FromString("<opc_helper.OPCServer not connected>");
	}
	return PyUnicode_FromFormat("<opc_helper.OPCServer connected, guid: %V, progid: %V>",
		self->m_szGuid, "None", self->m_szProgID, "None");
}


static PyObject *opch_IOPCServer_disconnect(opch_IOPCServer *self) {
	if (self) {
		if (self->m_srv) {
			self->m_srv->Release();
			self->m_srv = NULL;
		}
	}
	Py_RETURN_NONE;
}


static time_t opch_FileTimeToTimeT(const FILETIME *lpFileTime) {
	SYSTEMTIME st;
	FileTimeToSystemTime(lpFileTime, &st);
	struct tm ts;
	ts.tm_isdst = 0;
	ts.tm_wday = 0;
	ts.tm_yday = 0;
	// day of week
	// ts.tm_wday = (int)st.wDayOfWeek;
	// The values of the members tm_wday and tm_yday are ignored
	// date
	ts.tm_year = (int)st.wYear - 1900; // tm_year: Year (current year minus 1900)
	ts.tm_mon = (int)st.wMonth;
	ts.tm_mday = (int)st.wDay;
	// time
	ts.tm_hour = (int)st.wHour;
	ts.tm_min = (int)st.wMinute;
	ts.tm_sec = (int)st.wSecond;
	//
	return mktime(&ts);
}


static PyObject *opch_IOPCServer_get_status(opch_IOPCServer *self) {
	if (!self->m_srv) {
		PyErr_SetString(PyExc_RuntimeError, "Not connected!");
		return NULL;
	}

	// allocate return dict before all
	PyObject *ret_dict = PyDict_New();
	if (!ret_dict) {
		return NULL;
	}

	OPCSERVERSTATUS *pst;
	HRESULT hr = S_OK;
	hr = self->m_srv->GetStatus(&pst);
	if (FAILED(hr)) {
		PyErr_SetFromWindowsErr(hr);
		Py_DECREF(ret_dict);
		return NULL;
	}

	if (pst) {
		//
		// server times
		PyDict_SetItemString(ret_dict, "start_time",
			PyLong_FromLongLong(opch_FileTimeToTimeT(&(pst->ftStartTime))));
		PyDict_SetItemString(ret_dict, "current_time",
			PyLong_FromLongLong(opch_FileTimeToTimeT(&(pst->ftCurrentTime))));
		PyDict_SetItemString(ret_dict, "last_update_time",
			PyLong_FromLongLong(opch_FileTimeToTimeT(&(pst->ftLastUpdateTime))));
		//
		// server status
		switch(pst->dwServerState) {
		case OPC_STATUS_RUNNING:
			PyDict_SetItemString(ret_dict, "status", PyUnicode_FromString("OPC_STATUS_RUNNING"));
			break;
		case OPC_STATUS_FAILED:
			PyDict_SetItemString(ret_dict, "status", PyUnicode_FromString("OPC_STATUS_FAILED"));
			break;
		case OPC_STATUS_NOCONFIG:
			PyDict_SetItemString(ret_dict, "status", PyUnicode_FromString("OPC_STATUS_NOCONFIG"));
			break;
		case OPC_STATUS_SUSPENDED:
			PyDict_SetItemString(ret_dict, "status", PyUnicode_FromString("OPC_STATUS_SUSPENDED"));
			break;
		case OPC_STATUS_TEST:
			PyDict_SetItemString(ret_dict, "status", PyUnicode_FromString("OPC_STATUS_TEST"));
			break;
		case OPC_STATUS_COMM_FAULT:
			PyDict_SetItemString(ret_dict, "status", PyUnicode_FromString("OPC_STATUS_COMM_FAULT"));
			break;
		}
		//
		// group count
		PyDict_SetItemString(ret_dict, "group_count", PyLong_FromLong(pst->dwGroupCount));
		//
		// server version numbers
		char server_version_string[64];
		snprintf(server_version_string, sizeof(server_version_string)-1,
			"%d.%d.%d",
			(int)pst->wMajorVersion, (int)pst->wMinorVersion, (int)pst->wBuildNumber);
		PyDict_SetItemString(ret_dict, "server_version", PyUnicode_FromString(server_version_string));
		//
		// vendor info
		PyDict_SetItemString(ret_dict, "vendor_info", PyUnicode_FromString(""));
		if (pst->szVendorInfo) {
			PyDict_SetItemString(ret_dict, "vendor_info", PyUnicode_FromWideChar(pst->szVendorInfo, -1));
			CoTaskMemFree(pst->szVendorInfo);
		}
		CoTaskMemFree(pst);
	}

	return ret_dict;
}


// returns non-null if OK, NULL on exception thrown
static PyObject *opch_IOPCServer_browse_v3_recurse(IOPCBrowse *pBrowse, PyObject *list, Py_UNICODE *start_branch) {
	HRESULT hr = S_OK;
	WCHAR *pszContinuationPoint = NULL;
    BOOL bMoreElements = FALSE;
    DWORD dwCount = 0;
    OPCBROWSEELEMENT *pElements = NULL;
    DWORD dwPropertyCount = 0;
    DWORD dwPropertyIDs[2] = {1, 2};
    const int NO_RETURN_ITEMS_LIMIT = 0; // 0 means no limit

	// check start_branch
	if (start_branch == NULL) {
		// should not use NULL strings in COM
		start_branch = L""; // empty wide string
	}

	hr = pBrowse->Browse(start_branch,
		&pszContinuationPoint,
        NO_RETURN_ITEMS_LIMIT, // max number of elements to return
        OPC_BROWSE_FILTER_ALL, // OPC_BROWSE_FILTER_BRANCHES OPC_BROWSE_FILTER_ITEMS
        L"", // szElementNameFilter
		L"", // szVendorFilter
        TRUE, // bReturnAllProperties
		FALSE, // bReturnPropertyValues
        dwPropertyCount,
		dwPropertyIDs,
		&bMoreElements,
        &dwCount,       // [out], returned elements count
		&pElements );   // [out], returned elements

	if (SUCCEEDED(hr)) {
		unsigned i;
		for (i=0; i<dwCount; i++) {
			OPCBROWSEELEMENT *el = &pElements[i];
			BOOL hasChildren = FALSE;
			//
			PyObject *dict = PyDict_New();
			if (dict) {
				// only if dict was created
				PyObject *pyszItemID = PyUnicode_FromWideChar(el->szItemID, -1);
				PyObject *pyszName = PyUnicode_FromWideChar(el->szName, -1);
				if (el->dwFlagValue & OPC_BROWSE_HASCHILDREN) hasChildren = TRUE;
				//
				PyDict_SetItemString(dict, "name", pyszName);
				PyDict_SetItemString(dict, "item_id", pyszItemID);
				if (hasChildren) {
					PyObject *child_list = PyList_New(0);
					// recursively browse into this
					opch_IOPCServer_browse_v3_recurse(pBrowse, child_list, el->szItemID);
					// fill children
					PyDict_SetItemString(dict, "children", child_list);
				} else {
					Py_INCREF(Py_None);
					PyDict_SetItemString(dict, "children", Py_None);
				}
				//
				PyList_Append(list, dict);
			}
			//
			if (el->szItemID) CoTaskMemFree(el->szItemID);
			if (el->szName) CoTaskMemFree(el->szName);
		}

		if (pElements) CoTaskMemFree(pElements);
	} else {
		PySys_WriteStderr("\nIOPCBrowse::Browse() HRESULT = %d (0x%08X)\n", hr, hr);
		PyErr_SetFromWindowsErr(hr);
		return NULL;
	}
	return list;
}


static PyObject *opch_IOPCServer_browse_v3(opch_IOPCServer *self, Py_UNICODE *start_branch) {
	HRESULT hr = S_OK;
    IOPCBrowse *pBrowse = NULL;

	// create IOPCBrowse
	hr = self->m_srv->QueryInterface(IID_IOPCBrowse, (void **)&pBrowse);
	if (FAILED(hr)) {
		PyErr_SetFromWindowsErr(hr);
		return NULL;
	}

	if (start_branch == NULL) {
		// should not use NULL strings in COM
		start_branch = L""; // empty wide string
	}

	// create list that we will return
	PyObject *list = PyList_New(0);
	if (!list) {
		pBrowse->Release();
		return NULL;
	}

	PyObject *ret = opch_IOPCServer_browse_v3_recurse(pBrowse, list, start_branch);
	// we don't need IOPCBrowse at this point any more
	pBrowse->Release();
	pBrowse = NULL;

	// check recursive browser return value
	if (ret == NULL) {
		// some error happened inside
		Py_DECREF(list); // free list
		list = NULL;     // return error indicator
	}

	return list;
}


static PyObject *opch_IOPCServer_browse_v2_recurse(
	IOPCBrowseServerAddressSpace *pBrowse,
	PyObject *list,            // list to fill in the result
	Py_UNICODE *start_branch,  // branch name to start from (full ID) / empty str to browse from root
	BOOL use_flat)
{
	HRESULT hr = S_OK;
	IEnumString *pEnumString = NULL;

	// first check OPC server address space type
	// if it is flat, forcefully set use_flat = TRUE
	OPCNAMESPACETYPE nsType;
    pBrowse->QueryOrganization(&nsType);
	if (nsType == OPC_NS_FLAT)
		use_flat = TRUE;

	// check for NULL string
	if (start_branch == NULL) {
		// should not use NULL strings in COM
		start_branch = L""; // empty wide string
	}

	if (use_flat) {
        hr = pBrowse->BrowseOPCItemIDs(
			OPC_FLAT, // browseFilterType // this parameter is ignored for FLAT address space
			L"", // szFilterCriteria
			VT_EMPTY, // dataTypeFilter
			OPC_READABLE | OPC_WRITEABLE, // accessRightsFilter
			&pEnumString); // [out]
        if (FAILED(hr)) {
            PySys_WriteStderr("Failed BrowseOPCItemIDs(OPC_FLAT)");
			PyErr_SetFromWindowsErr(hr);
			return NULL;
        }
        ULONG fetched = 0;
        LPOLESTR pszName = NULL;
        while(pEnumString->Next(1, &pszName, &fetched) == S_OK) {
            if (pszName) {
				LPWSTR pszItemID = NULL;
				if (SUCCEEDED(pBrowse->GetItemID(pszName, &pszItemID))) {
					PyObject *dict = PyDict_New();
					if (dict) {
						PyDict_SetItemString(dict, "name", PyUnicode_FromWideChar(pszName, -1));
						PyDict_SetItemString(dict, "item_id", PyUnicode_FromWideChar(pszItemID, -1));
						Py_INCREF(Py_None);
						PyDict_SetItemString(dict, "children", Py_None);
						PyList_Append(list, dict);
					}
					CoTaskMemFree(pszItemID);
				}
				CoTaskMemFree(pszName);
			}
        }
        if (pEnumString)
			pEnumString->Release();
		return list;
	} // </if use_flat>

	// browse hierarchical namespace, as it normally goes
	// go to specified browse position
    hr = pBrowse->ChangeBrowsePosition(OPC_BROWSE_TO, start_branch);
	if (FAILED(hr)) {
		PySys_WriteStderr("Failed to change browse position to specified branch!\n");
		PyErr_SetFromWindowsErr(hr);
		return NULL;
	}

	//PySys_WriteStderr("In hierarchical browsing, changed browse position to %S\n", start_branch);

	// first enumerate all branches
    hr = pBrowse->BrowseOPCItemIDs(
		OPC_BRANCH, // browseFilterType
		L"", // szFilterCriteria
		VT_EMPTY, // dataTypeFilter
		OPC_READABLE | OPC_WRITEABLE, // accessRightsFilter
		&pEnumString); // [out]
    if (FAILED(hr)) {
        PySys_WriteStderr("Failed BrowseOPCItemIDs(OPC_BRANCH)");
		PyErr_SetFromWindowsErr(hr);
		return NULL;
    }

	ULONG fetched = 0;
    LPOLESTR pszName = NULL;
    while(pEnumString->Next(1, &pszName, &fetched) == S_OK) {
        if (pszName) {
			LPWSTR pszItemID = NULL;
			if (SUCCEEDED(pBrowse->GetItemID(pszName, &pszItemID))) {
				PyObject *dict = PyDict_New();
				if (dict) {
					PyDict_SetItemString(dict, "name", PyUnicode_FromWideChar(pszName, -1));
					PyDict_SetItemString(dict, "item_id", PyUnicode_FromWideChar(pszItemID, -1));
					
					// create a list for child items and set it
					PyObject *child_list = PyList_New(0);
					if (!child_list) {
						Py_DECREF(dict);
						CoTaskMemFree(pszItemID);
						CoTaskMemFree(pszName);
						pEnumString->Release();
						return NULL;
					}

					PyDict_SetItemString(dict, "children", child_list);
					PyList_Append(list, dict);
				}
				CoTaskMemFree(pszItemID);
			} else {
				PyObject *pyszName = PyUnicode_FromWideChar(pszName, -1);
				PySys_FormatStderr("Failed to get itemID for name [%V]\n", pyszName, "None");
				Py_DECREF(pyszName);
			}
			CoTaskMemFree(pszName);
		}
    }
    if (pEnumString)
		pEnumString->Release();
	pEnumString = NULL;

	// explicitly fill child items in list
	Py_ssize_t list_len = PyList_Size(list);
	Py_ssize_t list_index = 0;
	for (list_index = 0; list_index < list_len; list_index++) {
		PyObject *list_item = PyList_GetItem(list, list_index);
		if (list_item) {
			if (PyDict_Check(list_item)) {
				// this list item is dict.
				// we need to get item_id from it, and children
				// and then run recursively browse on this sub-branch
				// borrowed references - no need to decref
				PyObject *pyoItemID = PyDict_GetItemString(list_item, "item_id");
				PyObject *pyoChildren = PyDict_GetItemString(list_item, "children");
				if (pyoItemID && pyoChildren) {
					// itemID should be unicode
					// and children should be a list
					if (PyUnicode_Check(pyoItemID) && PyList_Check(pyoChildren)) {
						wchar_t *szItemID = PyUnicode_AsWideCharString(pyoItemID, NULL);
						if (szItemID) {
							// recurse into
							opch_IOPCServer_browse_v2_recurse(pBrowse, pyoChildren, szItemID, FALSE);
							PyMem_Free(szItemID);
						}
					}
				}
			}
		}
	}

	// then enumerate normal items
	// return to requested browse position
	hr = pBrowse->ChangeBrowsePosition(OPC_BROWSE_TO, start_branch);
	if (FAILED(hr)) {
        PySys_WriteStderr("Failed BrowseOPCItemIDs(OPC_LEAF)");
		PyErr_SetFromWindowsErr(hr);
		return NULL;
    }
	hr = pBrowse->BrowseOPCItemIDs(
		OPC_LEAF, // browseFilterType
		L"", // szFilterCriteria
		VT_EMPTY, // dataTypeFilter
		OPC_READABLE | OPC_WRITEABLE, // accessRightsFilter
		&pEnumString); // [out]
    if (FAILED(hr)) {
        PySys_WriteStderr("Failed BrowseOPCItemIDs(OPC_LEAF)");
		PyErr_SetFromWindowsErr(hr);
		return NULL;
    }

	while(pEnumString->Next(1, &pszName, &fetched) == S_OK) {
        if (pszName) {
			LPWSTR pszItemID = NULL;
			if (SUCCEEDED(pBrowse->GetItemID(pszName, &pszItemID))) {
				PyObject *dict = PyDict_New();
				if (dict) {
					PyDict_SetItemString(dict, "name", PyUnicode_FromWideChar(pszName, -1));
					PyDict_SetItemString(dict, "item_id", PyUnicode_FromWideChar(pszItemID, -1));
					// LEAF items always no children, and no recursion
					Py_INCREF(Py_None);
					PyDict_SetItemString(dict, "children", Py_None);
					PyList_Append(list, dict);
				}
				CoTaskMemFree(pszItemID);
			}
			CoTaskMemFree(pszName);
		}
    }
    if (pEnumString)
		pEnumString->Release();
	pEnumString = NULL;

	return list;
}


static PyObject *opch_IOPCServer_browse(opch_IOPCServer *self, PyObject *args, PyObject *kwds) {
	if (!self->m_srv) {
		PyErr_SetString(PyExc_RuntimeError, "Not connected!");
		return NULL;
	}

	BOOL use_flat = FALSE;
	Py_UNICODE *start_branch = NULL;

	static char *kwds_list[] = {"flat", "start_branch", NULL};

	if (!PyArg_ParseTupleAndKeywords(args, kwds, "|pZ", kwds_list, &use_flat, &start_branch)) {
		return NULL;
	}

	if (self->m_bSupportsV3) {
		// use_flat is ignored for v3 browse
		return opch_IOPCServer_browse_v3(self, start_branch);
	}

	// if we are here, browse server address space using OPCDA v2 methods

	IOPCBrowseServerAddressSpace *pBrowseSAS = NULL;
	HRESULT hr = self->m_srv->QueryInterface(IID_IOPCBrowseServerAddressSpace, (void **)&pBrowseSAS);
	if (FAILED(hr)) {
		PySys_WriteStderr("Failed to create IOPCBrowseServerAddressSpace (OPCDA v2) interface!");
		PyErr_SetFromWindowsErr(hr);
		return NULL;
	}

	PyObject *list = PyList_New(0);

	if (!opch_IOPCServer_browse_v2_recurse(pBrowseSAS, list, start_branch, use_flat)){
		// internal error, fail
		Py_DECREF(list);
		list = NULL;
	}

	pBrowseSAS->Release();
	pBrowseSAS = NULL;

	return list;
}


static BOOL opch_IOPCServer_add_group(opch_IOPCServer *self) {
	HRESULT hr = S_OK;
	Py_UNICODE *pszName = MY_OPC_GROUP_NAME;
	BOOL bActive = TRUE;
	DWORD dwUpdateRateMsec = 100;
	DWORD dwRevisedUpdateRate = 0;
	LONG lTimeBias = 0;
	FLOAT fPercentDeadband = 0.0f;

	OPCHANDLE hClientGroup = 1;

	LCID lcID = self->m_server_lang;  // defaults to MAKELANGID(LANG_NEUTRAL, SUBLANG_SYS_DEFAULT)

	hr = self->m_srv->AddGroup(pszName, bActive, dwUpdateRateMsec, hClientGroup, &lTimeBias, &fPercentDeadband, lcID,
		&self->m_hServerGroup, &dwRevisedUpdateRate, IID_IOPCGroupStateMgt, (IUnknown **)&self->m_groupMgt);

	//PySys_WriteStderr("self->m_srv->AddGroup() HRESULT = %d (0x%08X)\n", hr, hr);
	//PySys_WriteStderr("     hServerGroup = 0x%08X\n", self->m_hServerGroup);

	if (FAILED(hr)) {
		return FALSE;
	}

	// get IOPCItemMgt
	hr = self->m_groupMgt->QueryInterface(IID_IOPCItemMgt, (void **)&self->m_itemMgt);
	//PySys_WriteStderr("QueryInterface(IID_IOPCItemMgt) HRESULT = %d (0x%08X)\n", hr, hr);
	//PySys_WriteStderr("     self->m_itemMgt = 0x%08X\n", self->m_itemMgt);
	if (FAILED(hr)) {
		return FALSE;
	}

	// get IOPCSyncIO
	hr = self->m_groupMgt->QueryInterface(IID_IOPCSyncIO, (void **)&self->m_syncIO);
	//PySys_WriteStderr("QueryInterface(IID_IOPCSyncIO) HRESULT = %d (0x%08X)\n", hr, hr);
	//PySys_WriteStderr("     self->m_syncIO = 0x%08X\n", self->m_syncIO);
	if (FAILED(hr)) {
		return FALSE;
	}

	return TRUE;
}


static PyObject *opch_IOPCServer_get_opc_item(opch_IOPCServer *self, PyObject *args) {
	if (!self->m_srv) {
		PyErr_SetString(PyExc_RuntimeError, "Not connected!");
		return NULL;
	}

	if ((self->m_groupMgt == NULL) && (self->m_itemMgt == NULL)) {
		// create a group first
		BOOL ret = opch_IOPCServer_add_group(self);
		if (!ret) Py_RETURN_NONE;
	}

	Py_UNICODE *item_id = NULL;
	HRESULT hr = S_OK;

	if (!PyArg_ParseTuple(args, "u", &item_id)) {
		return NULL;
	}

	// first check maybe we already added this item to group
	struct opch_added_item *already_added = opch_IOPCServer_find_added_item(self, item_id);

	//if (already_added) {
	//	PySys_WriteStderr("Found existing already added item with serverHandle %d\n",
	//		already_added->hServerHandle);
	//}

	if (!already_added) {
		// add new
		OPCHANDLE hClientHandle = 1;
		OPCITEMDEF opcItem;
		memset(&opcItem, 0, sizeof(OPCITEMDEF));
		opcItem.bActive = TRUE;
		opcItem.szAccessPath = L"";
		opcItem.szItemID = item_id;
		opcItem.vtRequestedDataType = VT_EMPTY;
		opcItem.hClient = hClientHandle; // it can be any value and does not need to be unique
		// ^^ unique is reqired only for async r/w operations
		opcItem.dwBlobSize = 0;
		opcItem.pBlob = NULL;

		OPCITEMRESULT *pResult = NULL;
		HRESULT *phrErrors = NULL;
		hr = self->m_itemMgt->AddItems(1, &opcItem, &pResult, &phrErrors);
		// not interested in errors array
		if (phrErrors) {
			CoTaskMemFree(phrErrors);
			phrErrors = NULL;
		}
		if (SUCCEEDED(hr)) {
			if (pResult) {
				DWORD dwAccessRights = pResult->dwAccessRights; // OPC_READABLE | OPC_WRITEABLE;
				VARTYPE vtDataType = pResult->vtCanonicalDataType;
				OPCHANDLE hServerHandle = pResult->hServer;
				// For AddItems(), pBlob will alway be returned by servers
				// that support this feature
				if (pResult->pBlob) {
					CoTaskMemFree(pResult->pBlob);
					pResult->pBlob = NULL;
				}
				CoTaskMemFree(pResult);
				pResult = NULL;
				//
				//PySys_WriteStderr("added item [%ls] dataType=0x%08X, accessRights=0x%08X, server handle = %d\n",
				//	item_id, vtDataType, dwAccessRights, hServerHandle);
				// added item [Numeric._I4] dataType=0x00000003, accessRights=0x00000003, server handle = 1
				// add to list
				already_added = (struct opch_added_item *)malloc(sizeof(struct opch_added_item));
				if (!already_added) {
					PyErr_SetString(PyExc_MemoryError, "");
					return NULL;
				}
				already_added->dwAccessRights = dwAccessRights;
				already_added->hClientHandle = hClientHandle;
				already_added->hServerHandle = hServerHandle;
				already_added->vtDataType = vtDataType;
                already_added->tmTimeStamp = 0;
                already_added->wQuality = OPC_QUALITY_BAD;
				already_added->pszItemID = (wchar_t *)malloc((wcslen(item_id) + 1) * sizeof(wchar_t));
				if (already_added->pszItemID) {
					wcscpy(already_added->pszItemID, item_id);
				}
				already_added->next = NULL;
				// insert at the end of the list
				if (self->added_items == NULL) {
					self->added_items = already_added;
				} else {
					struct opch_added_item *cur = self->added_items;
					while (cur->next) cur = cur->next;
					cur->next = already_added;
				}
			}
		} else {
			PyErr_SetFromWindowsErr(hr);
			return NULL;
		}
	}

	// if we are here, either item was added new, or found existing,
	// and already_added points to item information

	OPCITEMSTATE *pItemValues = NULL;
	HRESULT *phrErrors = NULL;
	OPCHANDLE hServerHandle = already_added->hServerHandle;
	hr = self->m_syncIO->Read(OPC_DS_CACHE, 1, &hServerHandle, &pItemValues, &phrErrors);

	//PySys_WriteStderr("IOPCSyncIO::Read() HRESULT = %d (0x%08X)\n", hr, hr);

	PyObject *ret = NULL;

	if (SUCCEEDED(hr) && pItemValues) {
		//PySys_WriteStderr("SECCEDED, vt = 0x%08X\n", (unsigned)pItemValues->vDataValue.vt);
        // also store item timestamp?
        already_added->tmTimeStamp = opch_FileTimeToTimeT(&pItemValues->ftTimeStamp);
        already_added->wQuality = pItemValues->wQuality;
        //
		switch (pItemValues->vDataValue.vt) {
		case VT_EMPTY:
		case VT_NULL:
			Py_INCREF(Py_None);
			ret = Py_None;
			break;
		case VT_I1:
			ret = PyLong_FromLong(pItemValues->vDataValue.cVal);
			break;
		case VT_I2:
			ret = PyLong_FromLong(pItemValues->vDataValue.iVal);
			break;
		case VT_I4:
			ret = PyLong_FromLong(pItemValues->vDataValue.lVal);
			break;
		case VT_UI1:
			ret = PyLong_FromLong(pItemValues->vDataValue.bVal);
			break;
		case VT_UI2:
			ret = PyLong_FromLong(pItemValues->vDataValue.uiVal);
			break;
		case VT_UI4:
			ret = PyLong_FromLong(pItemValues->vDataValue.ulVal);
			break;
		case VT_R4:
			ret = PyFloat_FromDouble((double)pItemValues->vDataValue.fltVal);
			break;
		case VT_R8:
			ret = PyFloat_FromDouble(pItemValues->vDataValue.dblVal);
			break;
		case VT_INT:
			ret = PyLong_FromLong(pItemValues->vDataValue.intVal);
			break;
		case VT_UINT:
			ret = PyLong_FromLong((long)(pItemValues->vDataValue.uintVal));
			break;
		case VT_BSTR:
            if (pItemValues->vDataValue.bstrVal)
                ret = PyUnicode_FromWideChar(pItemValues->vDataValue.bstrVal, -1);
            else
                ret = PyUnicode_FromString("");
			break;
		case VT_BOOL:
			ret = PyBool_FromLong(pItemValues->vDataValue.boolVal);
			break;
		case VT_DATE:
		case VT_CY:
			VARIANT v2;
			VariantInit(&v2);
			hr = VariantChangeType(&v2, &pItemValues->vDataValue, 0, VT_BSTR);
			if (SUCCEEDED(hr)) {
				ret = PyUnicode_FromWideChar(v2.bstrVal, -1);
			}
			VariantClear(&v2);
			break;
		default:
			ret = NULL;
			PyErr_SetString(PyExc_TypeError, "Unhandled VARIANT data type!");
			break;
		}
	}

	if (phrErrors)
		CoTaskMemFree(phrErrors);
	if (pItemValues) {
		VariantClear(&pItemValues->vDataValue);
		CoTaskMemFree(pItemValues);
	}

	//PySys_WriteStderr("returning = %p\n", ret);
	return ret;
}


static PyObject *opch_IOPCServer_get_opc_item_info(opch_IOPCServer *self, PyObject *args) {
    if (!self->m_srv) {
        PyErr_SetString(PyExc_RuntimeError, "Not connected!");
        return NULL;
    }

    if ((self->m_groupMgt == NULL) && (self->m_itemMgt == NULL)) {
        PyErr_SetString(PyExc_RuntimeError, "No items were read from server, read this item first!");
        return NULL;
    }

    Py_UNICODE *item_id = NULL;
    struct opch_added_item *aitem = NULL;
    PyObject *ret_dict = NULL;

    if (!PyArg_ParseTuple(args, "u", &item_id)) {
        return NULL;
    }

    aitem = opch_IOPCServer_find_added_item(self, item_id);
    if (!aitem) {
        PyErr_SetString(PyExc_RuntimeError, "This item was never read from server, read this item first!");
        return NULL;
    }

    ret_dict = PyDict_New();
    if (ret_dict) {
        char szType[16] = {0};
        char szAT[16] = {0};
        char szQual[16] = {0};

        switch(aitem->vtDataType) {
        case VT_EMPTY: strcpy(szType, "VT_EMPTY"); break;
        case VT_NULL: strcpy(szType, "VT_NULL"); break;
        case VT_I1: strcpy(szType, "VT_I1"); break;
        case VT_I2: strcpy(szType, "VT_I2"); break;
        case VT_I4: strcpy(szType, "VT_I4"); break;
        case VT_UI1: strcpy(szType, "VT_UI1"); break;
        case VT_UI2: strcpy(szType, "VT_UI2"); break;
        case VT_UI4: strcpy(szType, "VT_UI4"); break;
        case VT_R4: strcpy(szType, "VT_R4"); break;
        case VT_R8: strcpy(szType, "VT_R8"); break;
        case VT_INT: strcpy(szType, "VT_INT"); break;
        case VT_UINT: strcpy(szType, "VT_UINT"); break;
        case VT_BSTR: strcpy(szType, "VT_BSTR"); break;
        case VT_BOOL: strcpy(szType, "VT_BOOL"); break;
        case VT_DATE: strcpy(szType, "VT_DATE"); break;
        case VT_CY: strcpy(szType, "VT_CY"); break;
        default: strcpy(szType, "UNKNOWN"); break;
        }

        if (aitem->dwAccessRights & OPC_READABLE) strcat(szAT, "read");
        if (aitem->dwAccessRights & OPC_WRITEABLE) strcat(szAT, "write");

        unsigned char qual = (unsigned char)(aitem->wQuality & 0x00FF);
        if ((qual & OPC_QUALITY_MASK) == OPC_QUALITY_BAD) strcpy(szQual, "BAD");
        if ((qual & OPC_QUALITY_MASK) == OPC_QUALITY_GOOD) strcpy(szQual, "GOOD");
        if ((qual & OPC_QUALITY_MASK) == OPC_QUALITY_UNCERTAIN) strcpy(szQual, "UNCERTAIN");

        PyDict_SetItemString(ret_dict, "type", PyUnicode_FromString(szType));
        PyDict_SetItemString(ret_dict, "access_rights", PyUnicode_FromString(szAT));
        PyDict_SetItemString(ret_dict, "timestamp", PyLong_FromLongLong((long long)aitem->tmTimeStamp));
        PyDict_SetItemString(ret_dict, "quality", PyUnicode_FromString(szQual));
    } else {
        PyErr_SetString(PyExc_MemoryError, "");
    }

    return ret_dict;
}


static PyObject *opch_IOPCServer_set_opc_item(opch_IOPCServer *self, PyObject *args) {
	Py_RETURN_NONE;
}


static PyObject *opch_IOPCServer_supports_v3(opch_IOPCServer *self) {
	if (!self->m_srv) {
		PyErr_SetString(PyExc_RuntimeError, "This object is not connected!");
		return NULL;
	}
	if (self->m_bSupportsV3) {
		Py_RETURN_TRUE;
	}
	Py_RETURN_FALSE;
}


static PyMemberDef opch_IOPCServer_members[] = {
	{"guid", T_OBJECT_EX, offsetof(opch_IOPCServer, m_szGuid), READONLY, "COM Object GUID"},
	{"progid", T_OBJECT_EX, offsetof(opch_IOPCServer, m_szProgID), READONLY, "COM Object ProgID"},
	{NULL} /* Sentinel */
};


static PyMethodDef opch_IOPCServer_methods[] = {
	{"supports_v3", (PyCFunction)opch_IOPCServer_supports_v3, METH_NOARGS,
		"supports_v3() -> bool\n"
		"Returns True, if this OPC Server supports OPCDA v3 spec."},
	{"disconnect", (PyCFunction)opch_IOPCServer_disconnect, METH_NOARGS,
		"disconnect() -> None\n"
		"Disconnects this OPC server (releases bound IOPCServer object)."},
	{"get_status", (PyCFunction)opch_IOPCServer_get_status, METH_NOARGS,
		"get_status() -> dict\n"
		"Query OPC server status. The returned dict will have fields:\n"
		"  'start_time', 'current_time', 'last_update_time' as unix timestamp in UTC;\n"
		"  'status' - string, in the form of 'OPC_STATUS_RUNNING'\n"
		"  'group_count' - int, for OPC debugging purposes\n"
		"  'server_version' - string\n"
		"  'vendor_info' - string\n"},
	{"browse", (PyCFunction)opch_IOPCServer_browse, METH_VARARGS | METH_KEYWORDS,
		"browse(flat: bool = False, start_branch: str = None) -> list\n"
		"List items in OPC server address space. Parameters:\n"
		"  flat: (bool) for OPCv2 servers only; treat server address \n"
		"                space as flat, return all items\n"
		"  start_branch: (str) starting node to enumerate from, if None\n"
		"               or empty string, begin enumeration from root node."},
	{"get_item", (PyCFunction)opch_IOPCServer_get_opc_item, METH_VARARGS,
		"get_item(item_id: str)\n"
		"Reads value of an OPC item.\n"},
	{"set_item", (PyCFunction)opch_IOPCServer_set_opc_item, METH_VARARGS,
		"set_item(item_id: str, value)\n"
		"Write value to an OPC item.\n"},
    {"get_item_info", (PyCFunction)opch_IOPCServer_get_opc_item_info, METH_VARARGS,
        "get_item_info(item_id: str)\n"
        "Returns dict with various information about an item:\n"
        " {\n"
        "   'type': 'VT_xxx', - item data type (VARIANT types)\n"
        "   'access_rights': 'read' / 'write' / 'readwrite'\n"
        "   'timestamp': unix UTC timestamp of last value update\n"
        "   'quality': 'BAD' / 'GOOD' / 'UNCERTAIN'\n"
        " }"},
	{NULL, NULL, 0, NULL} /* Sentinel */
};


PyTypeObject opch_IOPCServer_type = {
	PyVarObject_HEAD_INIT(NULL, 0)
	"opc_helper.OPCServer",     /* tp_name */
	sizeof(opch_IOPCServer),    /* tp_basicsize */
	0,                          /* tp_itemsize */
    (destructor)opch_IOPCServer_dealloc, /* tp_dealloc */
    0,                          /* tp_print */
    0,                          /* tp_getattr */
    0,                          /* tp_setattr */
    0,                          /* tp_reserved */
    (reprfunc)opch_IOPCServer_repr, /* tp_repr */
    0,                          /* tp_as_number */
    0,                          /* tp_as_sequence */
    0,                          /* tp_as_mapping */
    0,                          /* tp_hash  */
    0,                          /* tp_call */
    0,                          /* tp_str */
    0,                          /* tp_getattro */
    0,                          /* tp_setattro */
    0,                          /* tp_as_buffer */
    Py_TPFLAGS_DEFAULT | 
		Py_TPFLAGS_BASETYPE |
		Py_TPFLAGS_HAVE_GC,     /* tp_flags */
    PyDoc_STR("Represents OPC Server"),    /* tp_doc */
	(traverseproc)opch_IOPCServer_traverse, /* tp_traverse */
	(inquiry)opch_IOPCServer_clear, /* tp_clear */
    0,                          /* tp_richcompare */
    0,                          /* tp_weaklistoffset */
    0,                          /* tp_iter */
    0,                          /* tp_iternext */
    opch_IOPCServer_methods,    /* tp_methods */
    opch_IOPCServer_members,    /* tp_members */
    0,                          /* tp_getset */
    0,                          /* tp_base */
    0,                          /* tp_dict */
    0,                          /* tp_descr_get */
    0,                          /* tp_descr_set */
    0,                          /* tp_dictoffset */
    (initproc)opch_IOPCServer_init, /* tp_init */
    0,                          /* tp_alloc */
    opch_IOPCServer_new,        /* tp_new */
};


//======================================================================
// Factory function to create OPCServer object and connect to OPC server
// at the same time. This is the only way to create OPCServer object and
// connection to real OPC server in this API :)

PyObject *opch_opc_connect(PyObject *self, PyObject *args) {
	Py_UNICODE *szGuid = NULL;
	Py_UNICODE *szCompName = NULL;
	if (!PyArg_ParseTuple(args, "u|Z", &szGuid, &szCompName)) {
		return NULL;
	}

	GUID guid;
	HRESULT hr = S_OK;
	IOPCServer *pSrv = NULL;
	LPOLESTR lpszProgID = NULL;

	// convert guid from string to struct IID
	CLSIDFromString((LPCOLESTR)szGuid, &guid);
	ProgIDFromCLSID(guid, &lpszProgID);


	if ((szCompName == NULL) || (wcslen(szCompName) == 0)) {
		// create object on local computer
		hr = CoCreateInstance(guid, NULL, CLSCTX_ALL, IID_IOPCServer, (void **)&pSrv);
	} else {
		// create on remote computer
		COSERVERINFO coServerInfo;
		MULTI_QI mqi;

		memset(&coServerInfo, 0, sizeof(coServerInfo));
		coServerInfo.pAuthInfo = NULL; // default autorization
		coServerInfo.pwszName = szCompName;

		mqi.hr = S_OK;
		mqi.pIID = &IID_IOPCServer;
		mqi.pItf = NULL;

		hr = ::CoCreateInstanceEx(guid, NULL, CLSCTX_REMOTE_SERVER, &coServerInfo, 1, &mqi);
		if (SUCCEEDED(hr)) {
			pSrv = (IOPCServer *)mqi.pItf;
		}
	}

	if (FAILED(hr)) {
		PySys_WriteStderr("ERROR: Create COM object failed: HRESULT = %d (0x%08X)\n",
			(int)hr, (unsigned)hr);
		PyErr_SetFromWindowsErr(hr);
		return NULL;
	}

	opch_IOPCServer *opcsrv = (opch_IOPCServer *)opch_IOPCServer_new(&opch_IOPCServer_type, NULL, NULL);
	opcsrv->m_srv = pSrv;
	opcsrv->m_szGuid = PyUnicode_FromWideChar(szGuid, -1);
	if (lpszProgID) {
		opcsrv->m_szProgID = PyUnicode_FromWideChar(lpszProgID, -1);
		CoTaskMemFree(lpszProgID);
	}

	// try to get pointer to IOPCCommon interface
	IOPCCommon *pCmn = NULL;
	hr = opcsrv->m_srv->QueryInterface(IID_IOPCCommon, (void **)&pCmn);
	if (SUCCEEDED(hr)) {
		// usually this succeeds
		pCmn->SetClientName(L"PythonOPC_v1");
		pCmn->GetLocaleID(&opcsrv->m_server_lang);
		// PySys_WriteStderr("Got Locale ID = %d (0x%08X)\n", lcID, lcID);
		// PySys_WriteStderr("Primary = %x, sublang = %x\n", PRIMARYLANGID(lcID), SUBLANGID(lcID));
		// ^^  Got Locale ID = 2048 (0x00000800)
		// Primary = 0, sublang = 2
		pCmn->Release();
		pCmn = NULL;
	}

	// check if server supports OPCDA v3 version by querying for some v3-only interfaces
	IOPCBrowse *pBrowse = NULL; // check for IOPCBrowse
	hr = opcsrv->m_srv->QueryInterface(IID_IOPCBrowse, (void **)&pBrowse);
	if (FAILED(hr) || !pBrowse) {
		opcsrv->m_bSupportsV3 = FALSE;
	} else {
		if (pBrowse) {
			opcsrv->m_bSupportsV3 = TRUE;
			pBrowse->Release();
		}
	}

	return (PyObject *)opcsrv;
}
