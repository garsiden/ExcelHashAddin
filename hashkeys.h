// 
// Function prototypes
//

#pragma comment(linker, "/OUT:hashkeys.xll")

__declspec(dllexport)
signed long int WINAPI JenkinsHashKey (char *key);

__declspec(dllexport)
signed long int WINAPI HashKeyLittleEndian( char *key, signed long int initval);

