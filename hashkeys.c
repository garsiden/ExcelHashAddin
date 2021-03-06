#include <windows.h>
#include <stdlib.h>
#include <wchar.h>
#include <wctype.h>
#include <xlcall.h>
#include <stddef.h>
#include <stdint.h>
#include "hashkeys.h"
#include "jenkins.h"

// function definitions
int lpwstricmp(LPWSTR s, LPWSTR t);
LPXLOPER12 xlstring12(const XCHAR* lpstr);
XCHAR *byte_str(const XCHAR *);
XCHAR *byte_str_fromcs(const char *s);
uint32_t jenkins(char *key, size_t len);
uint32_t hashlittle(const void *key, size_t length, uint32_t initval);

//
// Syntax of the Register Command:
//      REGISTER(module_text, procedure, type_text, function_text, 
//               argument_text, macro_type, category, shortcut_text,
//               help_topic, function_help, argument_help1, argument_help2,...)
//
//
// functions will use only the first 11 arguments of 
// the Register function.
//
// This is a table of all the worksheet functions exported by this module.
// These functions are all registered (in xlAutoOpen) when you
// open the XLL. Before every string, leave a space for the
// byte count. The format of this table is the same as 
// arguments two through eleven of the REGISTER function.
// FUNCTION_ROWS define the number of rows in the table. The
// FUNCTION_COLS represents the number of columns in the table.
//
#define FUNCTION_ROWS 2
#define FUNCTION_COLS 11
#define TYPE_CMD 2
#define TYPE_FUN 1
#define FUNCTION_TEXT_COL 2
#define FUNCTION_PROC_COL 0
#define ADDIN_NAME L"Hash Keys XLL Add in"


static LPWSTR functions[FUNCTION_ROWS][FUNCTION_COLS] = {
     {L"JenkinsHashKey",                        // Procedure
      L"JF",                                    // char *, signed long int
      L"JenkinsHashKey",                        // function_text
      L"text",                                  // argument_text
      L"1",                                     // macro_type (2)
      L"Hash Keys",                             // category
      L"",                                      // shortcut_text
      L"",                                      // help_topic
      L"Jenkins one-at-a-time hash",            // function_help
      L"is the text string to hash",             // argument_help1
      L""                                       // argument_help2
     }
     ,
     {L"HashKeyLittleEndian",                   // Procedure
      L"JFN",                                   // char *, signed long int
      L"HashKeyLittleEndian",                   // function_text
      L"text,init_val",                         // argument_text
      L"1",                                     // macro_type (2)
      L"Hash Keys",                             // category
      L"",                                      // shortcut_text
      L"",                                      // help_topic
      L"Jenkins Little Endian hash",            // function_help
      L"is the text string to hash",            // argument_help1
      L"is an initial integer key value. "      // argument_help2
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
     XLOPER12 xDLL;				// name of this DLL
     int i;						// Loop indices
     int j;
     LPXLOPER12 pxRegArgs[FUNCTION_COLS];
     XLOPER12 xType;

     //
     // In the following block of code the name of the XLL is obtained by
     // calling xlGetName. This name is used as the first argument to the
     // REGISTER function to specify the name of the XLL. Next, the XLL loops
     // through the functions[] table, table registering each function in the
     // table using xlfRegister. 
     // Functions must be registered before you can add a menu item.
     // 

     Excel12(xlGetName, &xDLL, 0);
     xType.xltype = xltypeInt;
     xType.val.w = TYPE_FUN;

     for (i = 0; i < FUNCTION_ROWS; i++) {
          for (j = 0; j < FUNCTION_COLS; j++) {
               pxRegArgs[j] = xlstring12(functions[i][j]);
          }

          Excel12(xlfRegister, 0, 1 + FUNCTION_COLS, (LPXLOPER12) &xDLL,
                  pxRegArgs[0], pxRegArgs[1], pxRegArgs[2], pxRegArgs[3],
                  &xType, pxRegArgs[5], pxRegArgs[6], pxRegArgs[7],
                  pxRegArgs[8], pxRegArgs[9], pxRegArgs[10]);

          for (j = 0; j < FUNCTION_COLS; j++) {
               free(pxRegArgs[j]);
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
//      xlAutoClose is called by HASHKEYS.XLL by the function fExit. This
//      function is called when you exit Generic.
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
//      xlAutoClose does NOT have to unregister the functions that were
//      registered in xlAutoOpen. This is done automatically by Microsoft Excel
//      after xlAutoClose returns.
// 
//      xlAutoClose should return 1.
//
///***************************************************************************

__declspec(dllexport)
int WINAPI xlAutoClose(void)
{
     int i;
     LPXLOPER12 pxDefName;

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

     for (i = 0; i < FUNCTION_ROWS; i++) {
          pxDefName = xlstring12(functions[i][FUNCTION_TEXT_COL]);
          Excel12(xlfSetName, 0, 1, pxDefName);
          free(pxDefName);
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

     for (i = 1; i <= s[0]; i++) {
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
//      specify the type_text argument. If xlAutoRegister12 does not recognize
//      the function name, it should return a #VALUE! error. Otherwise, it
//      should whatever REGISTER returned.
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
     XLOPER12 xType;
     int i;
     int j;

     LPXLOPER12 pxRegArgs[FUNCTION_COLS];
     //
     // This block initializes xRegId to a #VALUE! error first. This is done in
     // case a function is not found to register. Next, the code loops through 
     // the functions in functions[] and uses lpwstricmp to determine if the 
     // current row in functions[] represents the function that needs to be 
     // registered. When it finds the proper row, the function is registered 
     // and the register ID is returned to Microsoft Excel. If no matching 
     // function is found, an xRegId is returned containing a #VALUE! error.
     //

     xRegId.xltype = xltypeErr;
     xRegId.val.err = xlerrValue;

     xType.xltype = xltypeInt;
     xType.val.w = TYPE_FUN;

     for (i = 0; i < FUNCTION_ROWS; i++) {
          if (!lpwstricmp(functions[i][0], pxName->val.str)) {
               for (j = 0; j < FUNCTION_COLS; j++) {
                    pxRegArgs[j] = xlstring12(functions[i][j]);
               }
               Excel12(xlfRegister, 0, 1 + FUNCTION_COLS,
                       (LPXLOPER12) &xDLL, pxRegArgs[0], pxRegArgs[1],
                       pxRegArgs[2], pxRegArgs[3], &xType, pxRegArgs[5],
                       pxRegArgs[6], pxRegArgs[7], pxRegArgs[8], pxRegArgs[9],
                       pxRegArgs[10]);

               for (j = 0; j < FUNCTION_COLS; j++) {
                    free(pxRegArgs[j]);
               }
               // Free oper returned by xl
               Excel12(xlFree, 0, 1, (LPXLOPER12) & xDLL);

               return (LPXLOPER12) &xRegId;
          }
     }

     return (LPXLOPER12) &xRegId;
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
     LPXLOPER12 pxMsg;
     XLOPER12 xInt;

     wsprintfW((LPWSTR)szBuf,
               L"Thank you for adding " ADDIN_NAME L"\nbuilt on %hs at %hs",
               __DATE__, __TIME__);

     // Display a dialog box indicating that the XLL was successfully added
     pxMsg = xlstring12(szBuf);
     xInt.xltype = xltypeInt;
     xInt.val.w = 2;

     Excel12(xlcAlert, 0, 2, pxMsg, &xInt);
     free(pxMsg);

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
     LPXLOPER12 pxMsg;
     XLOPER12 xInt;

     pxMsg = xlstring12(L"Thank you for removing " ADDIN_NAME);
     xInt.xltype = xltypeInt;
     xInt.val.w = 2;

     // Show a dialog box indicating that the XLL was successfully removed
     Excel12(xlcAlert, 0, 2, pxMsg, &xInt);
     free(pxMsg);

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
     static XLOPER12 xInfo;
     XLOPER12 xIntAction;
     XLOPER12 xIntType;

     //
     // This code coerces the passed-in value to an integer. This is how the
     // code determines what is being requested. If it receives a 1, 
     // it returns a string representing the long name. If it receives 
     // anything else, it returns a #VALUE! error.
     //

     xIntType.xltype = xltypeInt;
     xIntType.val.w = xltypeInt;
     Excel12(xlCoerce, &xIntAction, 2, xAction, (LPXLOPER12) &xIntType);

     if (xIntAction.val.w == 1) {
          xInfo.xltype = xltypeStr | xlbitDLLFree;
          xInfo.val.str = byte_str(ADDIN_NAME);
     } else {
          xInfo.xltype = xltypeErr;
          xInfo.val.err = xlerrValue;
     }

     //Word of caution - returning static XLOPER12s/XLOPER12s is not thread safe
     //for UDFs declared as thread safe, use alternate memory allocation
     //mechanisms
     return (LPXLOPER12) & xInfo;
}

__declspec(dllexport)
void WINAPI xlAutoFree12(LPXLOPER12 pxFree)
{
    if(pxFree->xltype & xltypeStr) {
        free(pxFree->val.str);
    }
}


///***************************************************************************
// fExit()
//
// Purpose:
//
//      This is a user-initiated routine to exit HASHKEYS.XLL You may be tempted
//      to call UNREGISTER("HASHKEYS.XLL") in this function. Don't do it! It
//      will have the effect of forcefully unregistering all of the functions in
//      this DLL, even if they are registered somewhere else! Instead,
//      unregister functions one at a time.
//
///***************************************************************************

__declspec(dllexport)
int WINAPI fExit(void)
{
     XLOPER12 xDLL;				// The name of this DLL //
     LPXLOPER12 pxFunc;            // The name of the function //
     XLOPER12 xRegId;			// The registration ID //
     int i;

     //
     // This code gets the DLL name. It then uses this along with information
     // from g_rgFuncs[] to obtain a REGISTER.ID() for each function. The
     // register ID is then used to unregister each function. Then the code
     // frees the DLL name and calls xlAutoClose.
     //

     Excel12(xlGetName, &xDLL, 0);

     for (i = 0; i < FUNCTION_ROWS; i++) {
          pxFunc = xlstring12(functions[i][FUNCTION_PROC_COL]);
          Excel12(xlfRegisterId, &xRegId, 2, (LPXLOPER12) &xDLL, pxFunc);
          free(pxFunc);
          Excel12(xlfUnregister, 0, 1, (LPXLOPER12) &xRegId);
     }

     Excel12(xlFree, 0, 1, (LPXLOPER12) & xDLL);

     return xlAutoClose();
}

XCHAR *byte_str(const XCHAR *lpstr)
{
    XCHAR *lps;
    size_t len;

    len = wcslen(lpstr);
    lps = (XCHAR *)malloc((len + 1) * 2);

    if (!lps) {
        return 0;
    }

    lps[0] = (BYTE)len;
    wmemcpy(lps + 1, lpstr, len);

    return lps;
}

XCHAR *byte_str_fromcs(const char *s)
{
    size_t len;
    XCHAR *lps;

    len = strlen(s);
    lps = (XCHAR *)malloc((len + 1) * 2);
    lps[0] = (BYTE)len;
    mbstowcs(lps + 1, s,len);

    return lps;
}

LPXLOPER12 xlstring12(const XCHAR *lpstr)
{
     LPXLOPER12 lpx;
     XCHAR *lps;
     int len;

     // get number of wchar values excluding null terminator
     len = wcslen(lpstr);

     lpx = (LPXLOPER12)malloc(sizeof(XLOPER12) + (len + 1) * 2);

     if (!lpx) {
          return 0;
     }

     lps = (XCHAR *) ((CHAR *)lpx + sizeof(XLOPER12));

     lps[0] = (BYTE)len;

     // can't wcscpy_s because of removal of null-termination
     wmemcpy(lps + 1, lpstr, len);
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

     for (hash = i = 0; i < len; ++i) {
          hash += key[i];
          hash += (hash << 10);
          hash ^= (hash >> 6);
     }
     hash += (hash << 3);
     hash ^= (hash >> 11);
     hash += (hash << 15);

     return hash;
}

signed long int WINAPI
HashKeyLittleEndian( char *key, signed long int initval)
{
     size_t len;

     len = strlen(key);
     signed long int val;

     val = (signed long int)hashlittle((void *)key, len, (uint32_t)initval);

     return val;
}

uint32_t hashlittle(const void *key, size_t length, uint32_t initval)
{
     uint32_t a, b, c;           /* internal state */
     union {
          const void *ptr;
          size_t i;
     }
     u;                          /* needed for Mac Powerbook G4 */

     /* Set up the internal state */
     a = b = c = 0xdeadbeef + ((uint32_t)length) + initval;

     u.ptr = key;
     if (HASH_LITTLE_ENDIAN && ((u.i & 0x3) == 0)) {
          const uint32_t *k = (const uint32_t *)key;  /* read 32-bit chunks */
          const uint8_t *k8;

          /* all but last block: aligned reads and affect 32 bits of (a,b,c) */
          while (length > 12) {
               a += k[0];
               b += k[1];
               c += k[2];
               mix(a, b, c);
               length -= 12;
               k += 3;
          }

          /*------------------------ handle the last (probably partial) block */
          /* 
           * "k[2]&0xffffff" actually reads beyond the end of the string, but
           * then masks off the part it's not allowed to read.  Because the
           * string is aligned, the masked-off tail is in the same word as the
           * rest of the string.  Every machine with memory protection I've seen
           * does it on word boundaries, so is OK with this.  But VALGRIND will
           * still catch it and complain.  The masking trick does make the hash
           * noticably faster for short strings (like English words).
           */

          switch (length) {
          case 12:
               c += k[2];
               b += k[1];
               a += k[0];
               break;
          case 11:
               c += k[2] & 0xffffff;
               b += k[1];
               a += k[0];
               break;
          case 10:
               c += k[2] & 0xffff;
               b += k[1];
               a += k[0];
               break;
          case 9:
               c += k[2] & 0xff;
               b += k[1];
               a += k[0];
               break;
          case 8:
               b += k[1];
               a += k[0];
               break;
          case 7:
               b += k[1] & 0xffffff;
               a += k[0];
               break;
          case 6:
               b += k[1] & 0xffff;
               a += k[0];
               break;
          case 5:
               b += k[1] & 0xff;
               a += k[0];
               break;
          case 4:
               a += k[0];
               break;
          case 3:
               a += k[0] & 0xffffff;
               break;
          case 2:
               a += k[0] & 0xffff;
               break;
          case 1:
               a += k[0] & 0xff;
               break;
          case 0:
               return c;   /* zero length strings require no mixing */
          }

     } else if (HASH_LITTLE_ENDIAN && ((u.i & 0x1) == 0)) {
          const uint16_t *k = (const uint16_t *)key;  /* read 16-bit chunks */
          const uint8_t *k8;

          /*---------- all but last block: aligned reads and different mixing */
          while (length > 12) {
               a += k[0] + (((uint32_t)k[1]) << 16);
               b += k[2] + (((uint32_t)k[3]) << 16);
               c += k[4] + (((uint32_t)k[5]) << 16);
               mix(a, b, c);
               length -= 12;
               k += 6;
          }

          /*------------------------ handle the last (probably partial) block */
          k8 = (const uint8_t *)k;
          switch (length) {
          case 12:
               c += k[4] + (((uint32_t)k[5]) << 16);
               b += k[2] + (((uint32_t)k[3]) << 16);
               a += k[0] + (((uint32_t)k[1]) << 16);
               break;
          case 11:
               c += ((uint32_t)k8[10]) << 16; /* fall through */
          case 10:
               c += k[4];
               b += k[2] + (((uint32_t)k[3]) << 16);
               a += k[0] + (((uint32_t)k[1]) << 16);
               break;
          case 9:
               c += k8[8];      /* fall through */
          case 8:
               b += k[2] + (((uint32_t)k[3]) << 16);
               a += k[0] + (((uint32_t)k[1]) << 16);
               break;
          case 7:
               b += ((uint32_t)k8[6]) << 16; /* fall through */
          case 6:
               b += k[2];
               a += k[0] + (((uint32_t)k[1]) << 16);
               break;
          case 5:
               b += k8[4];      /* fall through */
          case 4:
               a += k[0] + (((uint32_t)k[1]) << 16);
               break;
          case 3:
               a += ((uint32_t)k8[2]) << 16; /* fall through */
          case 2:
               a += k[0];
               break;
          case 1:
               a += k8[0];
               break;
          case 0:
               return c;        /* zero length requires no mixing */
          }

     } else {            /* need to read the key one byte at a time */
          const uint8_t *k = (const uint8_t *)key;

          /*---------- all but the last block: affect some 32 bits of (a,b,c) */
          while (length > 12) {
               a += k[0];
               a += ((uint32_t)k[1]) << 8;
               a += ((uint32_t)k[2]) << 16;
               a += ((uint32_t)k[3]) << 24;
               b += k[4];
               b += ((uint32_t)k[5]) << 8;
               b += ((uint32_t)k[6]) << 16;
               b += ((uint32_t)k[7]) << 24;
               c += k[8];
               c += ((uint32_t)k[9]) << 8;
               c += ((uint32_t)k[10]) << 16;
               c += ((uint32_t)k[11]) << 24;
               mix(a, b, c);
               length -= 12;
               k += 12;
          }

          /*--------------------------- last block: affect all 32 bits of (c) */
          switch (length) { /* all the case statements fall through */
          case 12:
               c += ((uint32_t)k[11]) << 24;
          case 11:
               c += ((uint32_t)k[10]) << 16;
          case 10:
               c += ((uint32_t)k[9]) << 8;
          case 9:
               c += k[8];
          case 8:
               b += ((uint32_t)k[7]) << 24;
          case 7:
               b += ((uint32_t)k[6]) << 16;
          case 6:
               b += ((uint32_t)k[5]) << 8;
          case 5:
               b += k[4];
          case 4:
               a += ((uint32_t)k[3]) << 24;
          case 3:
               a += ((uint32_t)k[2]) << 16;
          case 2:
               a += ((uint32_t)k[1]) << 8;
          case 1:
               a += k[0];
               break;
          case 0:
               return c;
          }
     }

     final(a, b, c);
     return c;
}
