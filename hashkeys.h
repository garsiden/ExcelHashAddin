// 
// Function prototypes
//

#pragma comment(linker, "/OUT:hashkeys.xll")

__declspec(dllexport) signed long int WINAPI JenkinsHashKey (char *key);
__declspec(dllexport) signed long int WINAPI HashKeyLittle( char *key, signed long int initval);
__declspec(dllexport) LPXLOPER12 WINAPI Func1(LPXLOPER12 x);
