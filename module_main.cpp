#include <Python.h>

#include "module.h"


// global var to store exception state
PyObject *g_opc_exception;


static PyObject *opch_version(PyObject *self, PyObject *args) {
	return PyUnicode_FromString("0.1");
}


static PyMethodDef module_methods[] = {
	{"initialize_com", opch_initialize_com, METH_NOARGS, "initialize_com()\n"
		"Call winapi CoInitializeEx(COINIT_MULTITHREADED)"},
	{"uninitialize_com", opch_uninitialize_com, METH_NOARGS, "uninitialize_com()\n"
		"Call winapi CoUninitiaize()"},
	{"opc_version", opch_version, METH_NOARGS, "opc_version() -> str\n"
		"Return library version string"},
	{"opc_connect", opch_opc_connect, METH_VARARGS, "opc_connect(guid: str, comp_name: str = None) -> OPCServer\n"
		"Connect to OPC Server with GUID guid on computer comp_name\n"
		"If comp_name is None or empty string, connect on local computer.\n"
		"Returns None on error, or OPCServer object on success."},
	{"opc_enum_query", opch_opc_enum_query, METH_VARARGS, "opc_enum_query(computer_name: str = None)\n"
		"Return list of available OPC servers.\n"
		"Computer name can be None or an empty string to query server on local computer.\n"
		"Every item will be a dict containing:\n"
		"    'progid': str\n"
		"    'desc': str\n"
		"    'guid': str\n"},
	{NULL, NULL, 0, NULL}
};

static struct PyModuleDef opcmodule = {
	PyModuleDef_HEAD_INIT,
	"opc_helper",
	"Support OPC Protocol for Python.",
	-1,
	module_methods
};


PyMODINIT_FUNC PyInit_opc_helper(void) {
    PyObject *mod = NULL;

	opch_IOPCServer_type.tp_new = PyType_GenericNew;
	if (PyType_Ready(&opch_IOPCServer_type) < 0)
		return NULL;

	mod = PyModule_Create(&opcmodule);
	if (mod == NULL)
		return NULL;

	g_opc_exception = PyErr_NewException("opc_helper.OPCException", NULL, NULL);
	Py_INCREF(g_opc_exception);
	PyModule_AddObject(mod, "OPCException", g_opc_exception);

	Py_INCREF(&opch_IOPCServer_type);
	PyModule_AddObject(mod, "OPCServer", (PyObject *)&opch_IOPCServer_type);

	return mod;
}
