#ifndef H_MODULE
#define H_MODULE

#include <Python.h>

// global var to store exception state
extern PyObject *g_opc_exception;

// type object for OPCServer
extern PyTypeObject opch_IOPCServer_type;


// module functions
extern PyObject *opch_initialize_com(PyObject *self, PyObject *args);
extern PyObject *opch_uninitialize_com(PyObject *self, PyObject *args);

extern PyObject *opch_opc_enum_query(PyObject *self, PyObject *args);

extern PyObject *opch_opc_connect(PyObject *self, PyObject *args);

#endif
