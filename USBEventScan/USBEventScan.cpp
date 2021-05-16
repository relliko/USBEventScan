#define _WIN32_DCOM
#include <iostream>
#include <string>
#include <sstream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h> // IWbemServices interface
#include <windows.h>

#pragma comment(lib, "wbemuuid.lib")


unsigned int POLLING_RATE = 1000; // in ms 


/**
* Prints the names of all the USB devices resulting from a 
* CIM_USBDevice query.
* Returns the number of devices scanned.
*/
int EnumQueryResults(IEnumWbemClassObject* pEnumerator, wstring* outData) {
    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    short n = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtPropID;
        VARIANT vtPropDescription;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Description", 0, &vtPropDescription, 0, 0);
        hr = pclsObj->Get(L"DeviceID", 0, &vtPropID, 0, 0);
        //wcout << "USB Device Name : " << vtProp.bstrVal << endl;
        wstringstream wss;
        wss << "Device Name: " << vtPropDescription.bstrVal << endl
            << "\t- ID: " << vtPropID.bstrVal << endl;
        *outData += wss.str();
        VariantClear(&vtPropID);
        VariantClear(&vtPropDescription);

        n++;

        pclsObj->Release();
    }
    pEnumerator->Release();

    return n;
}


int main(int argc, char** argv) {

    HRESULT hres;

    // Initialize COM. ------------------------------------------

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x"
            << hex << hres << endl;
        return 1;                  // Program has failed.
    }

    // Set general COM security levels --------------------------
    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );


    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;                    // Program has failed.
    }

    // Obtain the initial locator to WMI -------------------------

    IWbemLocator* pLoc = NULL;

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object."
            << " Err code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Connect to WMI through the IWbemLocator::ConnectServer method

    IWbemServices* pSvc = NULL;

    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
        NULL,                    // User name. NULL = current user
        NULL,                    // User password. NULL = current
        0,                       // Locale. NULL indicates current
        NULL,                    // Security flags.
        0,                       // Authority (for example, Kerberos)
        0,                       // Context object 
        &pSvc                    // pointer to IWbemServices proxy
    );

    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x"
            << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;


    // Set security levels on the proxy -------------------------

    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    short numDevices = -1;
    while (1) {
        // Use the IWbemServices pointer to make requests of WMI ----
        IEnumWbemClassObject* pEnumerator = NULL;
        hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("SELECT * FROM CIM_USBDevice"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator);

        if (FAILED(hres)) {
            cout << "Query for USB devices failed."
                << " Error code = 0x"
                << hex << hres << endl;
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            return 1;               // Program has failed.
        }

        wstring outData;

        // Print the data from the query
        short res = EnumQueryResults(pEnumerator, &outData);

        if (numDevices != res) {
            if (numDevices < res && numDevices >= 0) {
                cout << "Device plugged in." << endl;
            } else if (numDevices > res && numDevices >= 0) {
                cout << "Device unplugged." << endl;
            }
            else {
                cout << "Scanning USB ports..." << endl;
            }
            // device count changed, print the list
            wcout << outData << endl;
            cout << "Number of connected USB devices: " << res << endl;
            cout << "----------------------------------------" << endl << endl;
            cout << "Scanning USB ports..." << endl;

            numDevices = res;
        }

        Sleep(POLLING_RATE); // minimize cpu usage
    }

    // Cleanup
    // ========
    pSvc->Release();
    pLoc->Release();

    CoUninitialize();

    return 0;   // Program successfully completed.

}