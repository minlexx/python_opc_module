#include "module.h"
#include <Windows.h>

PyObject *opch_initialize_com(PyObject *self, PyObject *args) {
	HRESULT res = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	BOOL ok = SUCCEEDED(res);
	return PyBool_FromLong(ok);
}


PyObject *opch_uninitialize_com(PyObject *self, PyObject *args) {
	CoUninitialize();
	Py_RETURN_NONE;
}
