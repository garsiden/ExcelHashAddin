
//#define WIN32_DEFAULT_LIBS
#include <windows.h>
#include <stdlib.h>
#include <wchar.h>
#include <wctype.h>
#include <xlcall.h>
#include <stddef.h>
#include <stdint.h>
#include "hashkeys.h"

int lpwstricmp(LPWSTR s, LPWSTR t);
LPXLOPER12 byte_str12(const XCHAR* lpstr);
uint32_t jenkins(char *key, size_t len);
//
// Syntax of the Register Command:
//      REGISTER(module_text, procedure, type_text, function_text, 
//               argument_text, macro_type, category, shortcut_text,
//               help_topic, function_help, argument_help1, argument_help2,...)
//
//
// g_rgWorksheetFuncs will use only the first 11 arguments of 
// the Register function.
//
// This is a table of all the worksheet functions exported by this module.
// These functions are all registered (in xlAutoOpen) when you
// open the XLL. Before every string, leave a space for the
// byte count. The format of this table is the same as 
// arguments two through eleven of the REGISTER function.
// g_rgWorksheetFuncsRows define the number of rows in the table. The
// g_rgWorksheetFuncsCols represents the number of columns in the table.
//
#define g_rgWorksheetFuncsRows 2
#define g_rgWorksheetFuncsCols 10

static LPWSTR g_rgWorksheetFuncs
[g_rgWorksheetFuncsRows][g_rgWorksheetFuncsCols] =
{
	{ L"Func1",                                 // Procedure
		L"UU",                                  // type_text
		L"Func1",                               // function_text
		L"Arg",                                 // argument_text
		L"1",                                   // macro_type
		L"My Add-In",                      		// category
		L"",                                    // shortcut_text
		L"",                                    // help_topic
		L"Always returns the string 'Func1'",   // function_help
		L"Argument ignored"                     // argument_help1
	},

	{ L"JenkinsHashKey",                        // Procedure
		L"JF",                                  // type_text
		L"JenkinsHashKey",                      // function_text
		L"Arg",                                 // argument_text
		L"2",                                   // macro_type (2)
		L"My Add-In",                      		// category
		L"",                                    // shortcut_text
		L"",                                    // help_topic
		L"Jenkins one-at-atime hash key",   	// function_help
		L"Argument ignored"                     // argument_help1
	}	
};

///***************************************************************************
// xlAutoOpen()
//
// Purpose: 
//      Microsoft Excel call this function when the DLL is loaded.
//
//      Microsoft Excel uses xlAutoOpen to load XLL files.
//      When you open an XLL file, the only action
//      Microsoft Excel takes is to call the xlAutoOpen function.
//
//      More specifically, xlAutoOpen is called:
//
//       - when you open this XLL file from the File menu,
//       - when this XLL is in the XLSTART directory, and is
//         automatically opened when Microsoft Excel starts,
//       - when Microsoft Excel opens this XLL for any other reason, or
//       - when a macro calls REGISTER(), with only one argument, which is the
//         name of this XLL.
//
//      xlAutoOpen is also called by the Add-in Manager when you add this XLL 
//      as an add-in. The Add-in Manager first calls xlAutoAdd, then calls
//      REGISTER("EXAMPLE.XLL"), which in turn calls xlAutoOpen.
//
//      xlAutoOpen should:
//
//       - register all the functions you want to make available while this
//         XLL is open,
//
//       - add any menus or menu items that this XLL supports,
//
//       - perform any other initialization you need, and
//
//       - return 1 if successful, or return 0 if your XLL cannot be opened.
///***************************************************************************

__declspec(dllexport)
	int WINAPI xlAutoOpen(void)
{

	XLOPER12 xDLL;	// name of this DLL //
	int i;	// Loop indices //
	int j;
	LPXLOPER12 xRegArgs[g_rgWorksheetFuncsCols];

	//
	// In the following block of code the name of the XLL is obtained by
	// calling xlGetName. This name is used as the first argument to the
	// REGISTER function to specify the name of the XLL. Next, the XLL loops
	// through the g_rgWorksheetFuncs[] table, and the g_rgCommandFuncs[]
	// table registering each function in the table using xlfRegister. 
	// Functions must be registered before you can add a menu item.
	// xRegArgs[4]

	Excel12(xlGetName, &xDLL, 0);

	XLOPER12 xType;

	xType.xltype = xltypeInt;
	xType.val.w = 2;

	for (i = 0; i < g_rgWorksheetFuncsRows; i++)
	{
		for (j = 0; j < g_rgWorksheetFuncsCols; j++)
		{
			xRegArgs[j] = (LPXLOPER12)byte_str12(g_rgWorksheetFuncs[i][j]);
		}

		Excel12(xlfRegister, 0, 1 + g_rgWorksheetFuncsCols, (LPXLOPER12) & xDLL, xRegArgs[0], xRegArgs[1], xRegArgs[2], xRegArgs[3], &xType, xRegArgs[5], xRegArgs[6], xRegArgs[7], xRegArgs[8], xRegArgs[9]);

		for (j = 0; j < g_rgWorksheetFuncsCols; j++)
		{
			free(xRegArgs[j]);
		}
	}

	// Free the XLL filename //
	Excel12(xlFree, 0, 1, (LPXLOPER12) & xDLL);

	return 1;
}

///***************************************************************************
// xlAutoClose()
//
// Purpose: Microsoft Excel call this function when the DLL is unloaded.
//
//      xlAutoClose is called by Microsoft Excel:
//
//       - when you quit Microsoft Excel, or 
//       - when a macro sheet calls UNREGISTER(), giving a string argument
//         which is the name of this XLL.
//
//      xlAutoClose is called by the Add-in Manager when you remove this XLL from
//      the list of loaded add-ins. The Add-in Manager first calls xlAutoRemove,
//      then calls UNREGISTER("HASHKEYS.XLL"), which in turn calls xlAutoClose.
// 
//      xlAutoClose is called by HASHKEYS.XLL by the function fExit. This function
//      is called when you exit Generic.
// 
//      xlAutoClose should:
// 
//       - Remove any menus or menu items that were added in xlAutoOpen,
// 
//       - do any necessary global cleanup, and
// 
//       - delete any names that were added (names of exported functions, and 
//         so on). Remember that registering functions may cause names to 
//         be created.
// 
//      xlAutoClose does NOT have to unregister the functions that were registered
//      in xlAutoOpen. This is done automatically by Microsoft Excel after
//      xlAutoClose returns.
// 
//      xlAutoClose should return 1.
//
///***************************************************************************

__declspec(dllexport)
	int WINAPI xlAutoClose(void)
{
	int i;
	LPXLOPER12 def_name;

	//
	// This block first deletes all names added by xlAutoOpen or
	// xlAutoRegister12.
	//

	//
	// Due to a bug in Excel the following code to delete the defined names
	// does not work.  There is no way to delete these
	// names once they are Registered
	// The code is left in, in hopes that it will be
	// fixed in a future version.
	//

	for (i = 0; i < g_rgWorksheetFuncsRows; i++)
	{
		def_name = byte_str12(g_rgWorksheetFuncs[i][2]);
		Excel12(xlfSetName, 0, 1, def_name);
		free(def_name);
	}

	return 1;
}

///***************************************************************************
// lpwstricmp()
//
// Purpose: 
//
//      Compares a pascal string and a null-terminated C-string to see
//      if they are equal.  Method is case insensitive
//
// Parameters:
//
//      LPWSTR s    First string (null-terminated)
//      LPWSTR t    Second string (byte counted)
//
// Returns: 
//
//      int         0 if they are equal
//                  Nonzero otherwise
//
// Comments:
//
//      Unlike the usual string functions, lpwstricmp
//      doesn't care about collating sequence.
//
///***************************************************************************

int lpwstricmp(LPWSTR s, LPWSTR t)
{
	int i;

	if (wcslen(s) != *t)
		return 1;

	for (i = 1; i <= s[0]; i++)
	{
		if (towlower(s[i - 1]) != towlower(t[i]))
			return 1;
	}

	return 0;
}

///***************************************************************************
// xlAutoRegister12()
//
// Purpose:
//
//      This function is called by Microsoft Excel if a macro sheet tries to
//      register a function without specifying the type_text argument. If that
//      happens, Microsoft Excel calls xlAutoRegister12, passing the name of the
//      function that the user tried to register. xlAutoRegister12 should use the
//      normal REGISTER function to register the function, only this time it must
//      specify the type_text argument. If xlAutoRegister12 does not recognize the
//      function name, it should return a #VALUE! error. Otherwise, it should
//      return whatever REGISTER returned.
//
// Parameters:
//
//      LPXLOPER12 pxName   xltypeStr containing the
//                          name of the function
//                          to be registered. This is not
//                          case sensitive.
//
// Returns: 
//
//      LPXLOPER12          xltypeNum containing the result
//                          of registering the function,
//                          or xltypeErr containing #VALUE!
//                          if the function could not be
//                          registered.
///***************************************************************************

__declspec(dllexport)
	LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 xRegId;
	XLOPER12 xDLL;
	int i;
	int j;

	LPXLOPER12 xRegArgs[g_rgWorksheetFuncsCols];
	//
	// This block initializes xRegId to a #VALUE! error first. This is done in
	// case a function is not found to register. Next, the code loops through 
	// the functions in g_rgFuncs[] and uses lpwstricmp to determine if the 
	// current row in g_rgFuncs[] represents the function that needs to be 
	// registered. When it finds the proper row, the function is registered 
	// and the register ID is returned to Microsoft Excel. If no matching 
	// function is found, an xRegId is returned containing a #VALUE! error.
	// xRegArgs[4],

	xRegId.xltype = xltypeErr;
	xRegId.val.err = xlerrValue;

	XLOPER12 xType;

	xType.xltype = xltypeInt;
	xType.val.w = 2;

	for (i = 0; i < g_rgWorksheetFuncsRows; i++)
	{
		if (!lpwstricmp(g_rgWorksheetFuncs[i][0], pxName->val.str))
		{
			for (j = 0; j < g_rgWorksheetFuncsCols; j++)
			{
				xRegArgs[j] = byte_str12(g_rgWorksheetFuncs[i][j]);
			}
			Excel12(xlfRegister, 0, 1 + g_rgWorksheetFuncsCols, (LPXLOPER12) & xDLL, xRegArgs[0], xRegArgs[1], xRegArgs[2], xRegArgs[3], &xType, xRegArgs[5], xRegArgs[6], xRegArgs[7], xRegArgs[8], xRegArgs[9]);

			for (j = 0; j < g_rgWorksheetFuncsCols; j++)
			{
				free(xRegArgs[j]);
			}
			/// Free oper returned by xl //
			Excel12(xlFree, 0, 1, (LPXLOPER12) & xDLL);

			return (LPXLOPER12) & xRegId;
		}
	}

	return (LPXLOPER12) & xRegId;
}

///***************************************************************************
// xlAutoAdd()
//
// Purpose:
//
//      This function is called by the Add-in Manager only. When you add a
//      DLL to the list of active add-ins, the Add-in Manager calls xlAutoAdd()
//      and then opens the XLL, which in turn calls xlAutoOpen.
//
///***************************************************************************

__declspec(dllexport)
	int WINAPI xlAutoAdd(void)
{
	XCHAR szBuf[255];
	LPXLOPER12 msg;
	XLOPER12 xInt;

	wsprintfW((LPWSTR)szBuf, L"Thank you for adding ExcelAddin.XLL\n " L"built on %hs at %hs", __DATE__, __TIME__);

	// Display a dialog box indicating that the XLL was successfully added //
	msg = byte_str12(szBuf);
	xInt.xltype = xltypeInt;
	xInt.val.w = 2;

	Excel12(xlcAlert, 0, 2, msg, &xInt);
	free(msg);

	return 1;
}

///***************************************************************************
// xlAutoRemove()
//
// Purpose:
//
//      This function is called by the Add-in Manager only. When you remove
//      an XLL from the list of active add-ins, the Add-in Manager calls
//      xlAutoRemove() and then UNREGISTER("HASHKEYS.XLL").
//   
//      You can use this function to perform any special tasks that need to be
//      performed when you remove the XLL from the Add-in Manager's list
//      of active add-ins. For example, you may want to delete an
//      initialization file when the XLL is removed from the list.
//
///***************************************************************************

__declspec(dllexport)
	int WINAPI xlAutoRemove(void)
{
	LPXLOPER12 msg;
	XLOPER12 xInt;

	msg = byte_str12(L"Thank you for removing ExcelAddin.XLL!");
	xInt.xltype = xltypeInt;
	xInt.val.w = 2;

	// Show a dialog box indicating that the XLL was successfully removed //
	Excel12(xlcAlert, 0, 2, msg, &xInt);
	free(msg);

	return 1;
}

///***************************************************************************
// xlAddInManagerInfo12()
//
// Purpose:
//
//      This function is called by the Add-in Manager to find the long name
//      of the add-in. If xAction = 1, this function should return a string
//      containing the long name of this XLL, which the Add-in Manager will use
//      to describe this XLL. If xAction = 2 or 3, this function should return
//      #VALUE!.
//
// Parameters:
//
//      LPXLOPER12 xAction  What information you want. One of:
//                            1 = the long name of the
//                                add-in
//                            2 = reserved
//                            3 = reserved
//
// Returns: 
//
//      LPXLOPER12          The long name or #VALUE!.
//
///***************************************************************************

__declspec(dllexport)
	LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	static XLOPER12 xInfo, xIntAction, xIntType;

	//
	// This code coerces the passed-in value to an integer. This is how the
	// code determines what is being requested. If it receives a 1, 
	// it returns a string representing the long name. If it receives 
	// anything else, it returns a #VALUE! error.
	//

	xIntType.xltype = xltypeInt;
	xIntType.val.w = xltypeInt;
	Excel12(xlCoerce, &xIntAction, 2, xAction, (LPXLOPER12) & xIntType);

	if (xIntAction.val.w == 1)
	{
		xInfo.xltype = xltypeStr;
		xInfo.val.str = L"\025ExcelAddin Standalone DLL";
	}
	else
	{
		xInfo.xltype = xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPER12s/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms
	return (LPXLOPER12) & xInfo;
}

///***************************************************************************
// Func1()
//
// Purpose:
//
//      This is a typical user-defined function provided by an XLL.
//
// Parameters:
//
//      LPXLOPER12 x    (Ignored)
//
// Returns: 
//
//      LPXLOPER12      Always the string "Func1"
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

LPXLOPER12 WINAPI Func1(LPXLOPER12 x)
{
	static XLOPER12 xResult;

	//
	// This function demonstrates returning a string value. The return
	// type is set to a string and filled with the name of the function.
	//

	xResult.xltype = xltypeStr;
	xResult.val.str = L"\005Func1";

	//Word of caution - returning static XLOPER12s/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms
	return (LPXLOPER12) & xResult;
}

///***************************************************************************
// fExit()
//
// Purpose:
//
//      This is a user-initiated routine to exit HASHKEYS.XLL You may be tempted to
//      simply call UNREGISTER("HASHKEYS.XLL") in this function. Don't do it! It
//      will have the effect of forcefully unregistering all of the functions in
//      this DLL, even if they are registered somewhere else! Instead, unregister
//      the functions one at a time.
//
///***************************************************************************

__declspec(dllexport)
	int WINAPI fExit(void)
{
	XLOPER12 xDLL,	// The name of this DLL //
	         xFunc,	// The name of the function //
	         xRegId;	// The registration ID //
	int i;

	//
	// This code gets the DLL name. It then uses this along with information
	// from g_rgFuncs[] to obtain a REGISTER.ID() for each function. The
	// register ID is then used to unregister each function. Then the code
	// frees the DLL name and calls xlAutoClose.
	//

	// Make xFunc a string //
	xFunc.xltype = xltypeStr;

	Excel12(xlGetName, &xDLL, 0);

	for (i = 0; i < g_rgWorksheetFuncsRows; i++)
	{
		xFunc.val.str = (LPWSTR) (g_rgWorksheetFuncs[i][0]);
		Excel12(xlfRegisterId, &xRegId, 2, (LPXLOPER12) & xDLL, (LPXLOPER12) & xFunc);
		Excel12(xlfUnregister, 0, 1, (LPXLOPER12) & xRegId);
	}

	Excel12(xlFree, 0, 1, (LPXLOPER12) & xDLL);

	return xlAutoClose();
}

LPXLOPER12 byte_str12(const XCHAR *lpstr)
{
	LPXLOPER12 lpx;
	XCHAR *lps;
	int len;

	// get number of wchar values excluding null terminator
	len = lstrlenW(lpstr);

	lpx = (LPXLOPER12)malloc(sizeof(XLOPER12) + (len + 1) * 2);

	if (!lpx)
	{
		return 0;
	}

	lps = (XCHAR *) ((CHAR *)lpx + sizeof(XLOPER12));

	lps[0] = (BYTE)len;

	// can't wcscpy_s because of removal of null-termination
	wmemcpy_s(lps + 1, len + 1, lpstr, len);
	lpx->xltype = xltypeStr;
	lpx->val.str = lps;

	return lpx;
}

signed long int WINAPI JenkinsHashKey(char *key)
{
	size_t len;

	len = strlen(key);
	signed long int val;

	val = (signed long int)jenkins(key, len);

	return val;
}

uint32_t jenkins(char *key, size_t len)
{
	uint32_t hash, i;

	for (hash = i = 0; i < len; ++i)
	{
		hash += key[i];
		hash += (hash << 10);
		hash ^= (hash >> 6);
	}
	hash += (hash << 3);
	hash ^= (hash >> 11);
	hash += (hash << 15);

	return hash;
}
