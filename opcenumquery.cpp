#include <Python.h>
#include <Windows.h>

#include "module.h"
#include "opc/opccomn.h"
#include "opc/OpcEnum.h"
#include "opc/opcda.h"

//==============================================================================
// helpers for enum opc servers

static PyObject *browse_helper1(IOPCServerList *sl) {
    IEnumGUID *pEnum = NULL;
    ULONG nIMP = 3;
    CATID cIMP[3];
    cIMP[0] = CATID_OPCDAServer10;
    cIMP[1] = CATID_OPCDAServer20;
    cIMP[2] = CATID_OPCDAServer30;

    HRESULT hr = sl->EnumClassesOfCategories(nIMP, cIMP, 0, NULL, &pEnum);
    if (FAILED(hr) || !pEnum) {
        PyErr_SetString(g_opc_exception, "IOPCServerList::EnumClassesOfCategories() failed!");
        return NULL;
    }

	PyObject *list = PyList_New(0);
    GUID guid;
    ULONG ulFetched = 0;
	//
    while(pEnum->Next(1, &guid, &ulFetched) == S_OK) {
        LPOLESTR lpszProgID = NULL;
        LPOLESTR lpszUserType = NULL;
        hr = sl->GetClassDetails(guid, &lpszProgID, &lpszUserType);
        if (SUCCEEDED(hr)) {
			PyObject *dict = PyDict_New();
			Py_UNICODE str_guid[64] = {0};
			StringFromGUID2(guid, str_guid, sizeof(str_guid)/sizeof(str_guid[0])-1);
			//
			PyDict_SetItemString(dict, "progid", PyUnicode_FromWideChar(lpszProgID, -1));
			PyDict_SetItemString(dict, "desc", PyUnicode_FromWideChar(lpszUserType, -1));
            PyDict_SetItemString(dict, "guid", PyUnicode_FromWideChar(str_guid, -1));
            PyList_Append(list, dict);
        }
        if (lpszProgID) CoTaskMemFree(lpszProgID);
        if (lpszUserType) CoTaskMemFree(lpszUserType);
    }
    pEnum->Release();
    return list;
}


static PyObject *browse_helper2(IOPCServerList2 *sl) {
    IOPCEnumGUID *pEnum = NULL;
    ULONG nIMP = 3;
    CATID cIMP[3];
    cIMP[0] = CATID_OPCDAServer10;
    cIMP[1] = CATID_OPCDAServer20;
    cIMP[2] = CATID_OPCDAServer30;

    HRESULT hr = sl->EnumClassesOfCategories(nIMP, cIMP, 0, NULL, &pEnum);
    if (FAILED(hr) || !pEnum) {
        PyErr_SetString(g_opc_exception, "IOPCServerList2::EnumClassesOfCategories() failed!");
        return NULL;
    }

	PyObject *list = PyList_New(0);
    GUID guid;
    ULONG ulFetched = 0;

    while(pEnum->Next(1, &guid, &ulFetched) == S_OK) {
        LPOLESTR lpszProgID = NULL;
        LPOLESTR lpszUserType = NULL;
        LPOLESTR lpszVerIndProgID = NULL;
        hr = sl->GetClassDetails(guid, &lpszProgID, &lpszUserType, &lpszVerIndProgID);
        if (SUCCEEDED(hr)) {
			Py_UNICODE str_guid[64] = {0};
			PyObject *dict = PyDict_New();
			StringFromGUID2(guid, str_guid, sizeof(str_guid)/sizeof(str_guid[0])-1);
			//
            PyDict_SetItemString(dict, "progid", PyUnicode_FromWideChar(lpszProgID, -1));
			PyDict_SetItemString(dict, "desc",   PyUnicode_FromWideChar(lpszUserType, -1));
            PyDict_SetItemString(dict, "guid",   PyUnicode_FromWideChar(str_guid, -1));
            PyList_Append(list, dict);
        }
        if (lpszProgID) CoTaskMemFree(lpszProgID);
        if (lpszUserType) CoTaskMemFree(lpszUserType);
        if (lpszVerIndProgID) CoTaskMemFree(lpszVerIndProgID);
    }
    pEnum->Release();
    return list;
}


//==============================================================================


PyObject *opch_opc_enum_query(PyObject *self, PyObject *args) {
	Py_UNICODE *wstr_server = NULL;

	if (!PyArg_ParseTuple(args, "|Z", &wstr_server)) {
		return NULL;
	}

	IOPCServerList  *sl1 = NULL;
    IOPCServerList2 *sl2 = NULL;
    HRESULT hr = S_OK;
	PyObject *ret = NULL;

	if (wstr_server == NULL || wcslen(wstr_server) == 0) {
		// creating on local computer
		hr = ::CoCreateInstance(CLSID_OpcServerList, NULL,
                                CLSCTX_ALL | CLSCTX_REMOTE_SERVER,
                                IID_IOPCServerList2, (void **)&sl2);
        if (FAILED(hr)) {
            hr = ::CoCreateInstance(CLSID_OpcServerList, NULL,
                                    CLSCTX_ALL | CLSCTX_REMOTE_SERVER,
                                    IID_IOPCServerList, (void **)&sl1);
            if (FAILED(hr)) {
				PySys_WriteStderr("Failed to create OPC Server Enumerator, "
					"please make sure OPC Base Components are installed!");
				PyErr_SetFromWindowsErr(hr);
                return NULL;
            }

            ret = browse_helper1(sl1);
            sl1->Release();
            return ret;
        }

        ret = browse_helper2(sl2);
        sl2->Release();
	} else {
		// create enum on remote computer
		COSERVERINFO coServerInfo;
        memset(&coServerInfo, 0, sizeof(coServerInfo));
        coServerInfo.pAuthInfo = NULL; // default autorization
        coServerInfo.pwszName = wstr_server;

        MULTI_QI mqi;
        mqi.hr = S_OK;
        mqi.pIID = &IID_IOPCServerList2;
        mqi.pItf = NULL;
        hr = ::CoCreateInstanceEx(CLSID_OpcServerList, NULL, CLSCTX_REMOTE_SERVER, &coServerInfo, 1, &mqi);

        if (FAILED(hr)) {
            memset(&coServerInfo, 0, sizeof(coServerInfo));
            coServerInfo.pAuthInfo = NULL; // default autorization
            coServerInfo.pwszName = wstr_server;

            mqi.hr = S_OK;
            mqi.pIID = &IID_IOPCServerList;
            mqi.pItf = NULL;
            hr = ::CoCreateInstanceEx(CLSID_OpcServerList, NULL, CLSCTX_REMOTE_SERVER, &coServerInfo, 1, &mqi);

            if (FAILED(hr)) {
				PySys_WriteStderr("Failed to create OPC Server Enumerator on remote computer!\n");
				PyErr_SetFromWindowsErr(hr);
                return NULL;
            }

            sl1 = (IOPCServerList *)mqi.pItf;
            ret = browse_helper1(sl1);
            sl1->Release();

            return ret;
        }

        sl2 = (IOPCServerList2 *)mqi.pItf;
        ret = browse_helper2(sl2);
        sl2->Release();
	}
	return ret;
}
