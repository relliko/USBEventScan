#define _WIN32_DCOM
#include <iostream>
#include <string>
#include <sstream>
#include <algorithm>
#include <map>
#include <vector>
#include <comdef.h>
#include <Wbemidl.h> // IWbemServices interface
#include <windows.h>

using namespace std;

#pragma comment(lib, "wbemuuid.lib")

constexpr unsigned int POLLING_RATE = 250; // in ms 
constexpr short GREEN  = 2;
constexpr short RED    = 4;
constexpr short WHITE  = 15;

HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE); // Allows us to pick output colors

// Stores metadata about a USB device on the system
struct USBDevice_t {
    wstring description;
    wstring id;
};

// Function definitions
vector<wstring> keySetDifference(map<wstring, USBDevice_t> m1, map<wstring, USBDevice_t> m2);
void CoutRed(wstring s);
void CoutGreen(wstring s);
int InitializeCOM();
IWbemLocator* InitializeWMILocator();
IWbemServices* ConnectToWMI(IWbemLocator* pLoc);
IEnumWbemClassObject* QueryUSBDevices(IWbemServices* pSvc, IWbemLocator* pLoc);
int EnumQueryResults(IEnumWbemClassObject* pEnumerator, map<wstring, USBDevice_t>& umap,
    wstring* outData);


int main(int argc, char** argv) {
    // Initialize COM. ------------------------------------------
    InitializeCOM();
    // Obtain the initial locator to WMI -------------------------
    IWbemLocator* pLoc = InitializeWMILocator();
    // Connect to WMI through the IWbemLocator::ConnectServer method
    // 
    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    IWbemServices* pSvc = ConnectToWMI(pLoc);

    short numDevices = -1;
    unsigned int totalChanges = 0; // number of additions or removals 
    map<wstring, USBDevice_t> deviceMap; // The map of devices existing on the system

    // Main program loop
    while (1) {
        // Use the IWbemServices pointer to make requests of WMI ----
        IEnumWbemClassObject* pEnumerator = QueryUSBDevices(pSvc, pLoc);

        map<wstring, USBDevice_t> tempMap;
        wstring outData;

        // Enumerate through the data from the query
        short res = EnumQueryResults(pEnumerator, tempMap, &outData);
        // TODO: Compare temp map and the device map to isolate changes 


        if (numDevices != res) {
            system("CLS"); // clear command prompt to update it 

            // Isolate the name of the device that changed
            vector<wstring> diffVector = keySetDifference(deviceMap, tempMap);
            
            // device count changed, print the list
            wcout << outData << endl;

            if (numDevices < res && numDevices >= 0) {
                totalChanges++;
                wstring s = L"Device added: " + diffVector[0];
                CoutGreen(s);
            }
            else if (numDevices > res && numDevices >= 0) {
                totalChanges++;
                wstring s = L"Device removed: " + diffVector[0];
                CoutRed(s);
            }
            else {
                // This else only runs on first iteration of while loop
            }

            cout << "Number of connected USB devices: " << res << endl;
            cout << "Changes observed: " << totalChanges << endl;
            cout << "----------------------------------------" << endl << endl;
            cout << "Scanning USB ports..." << endl;

            deviceMap = tempMap;
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

/**
* Enumerates through all the USB devices resulting from a
* CIM_USBDevice query and adds them to a a map.
* Returns the number of devices resulting from the query.
*/
int EnumQueryResults(IEnumWbemClassObject* pEnumerator,
    map<wstring, USBDevice_t>& umap, wstring* outData) {

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    short n = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if (0 == uReturn) {
            break;
        }

        VARIANT vtPropID;
        VARIANT vtPropDescription;

        // Get the values of properties
        hr = pclsObj->Get(L"Description", 0, &vtPropDescription, 0, 0);
        hr = pclsObj->Get(L"DeviceID", 0, &vtPropID, 0, 0);

        wstringstream wss;
        wss << "Device Name: " << vtPropDescription.bstrVal << endl
            << "\t- ID: " << vtPropID.bstrVal << endl;
        // Add to the map passed in to this function 
        USBDevice_t device;
        device.description = vtPropDescription.bstrVal;
        device.id = vtPropID.bstrVal;
        umap[device.id] = device;
        *outData += wss.str();
        VariantClear(&vtPropID);
        VariantClear(&vtPropDescription);
        pclsObj->Release();

        n++;
    }
    pEnumerator->Release();

    return n;
}

// Isolates any wstring keys which do not exist in either m1 or m2 and returns them in a vector.
vector<wstring> keySetDifference(map<wstring, USBDevice_t> m1,
    map<wstring, USBDevice_t> m2) {
    vector<wstring> diffVector;
    vector<wstring> keysD;
    keysD.reserve(m1.size());
    vector<wstring> keysT;
    keysT.reserve(m2.size());
    for (auto& it : m1) {
        keysD.push_back(it.first);
    }
    for (auto& it : m2) {
        keysT.push_back(it.first);
    }
    set_symmetric_difference(keysD.begin(), keysD.end(),
        keysT.begin(), keysT.end(), back_inserter(diffVector));
    return diffVector;
}

// Prints a wstring in red text
void CoutRed(wstring s) {
    SetConsoleTextAttribute(hConsole, RED);
    wcout << s << endl;
    SetConsoleTextAttribute(hConsole, WHITE);
}

// Prints a wstring in green text
void CoutGreen(wstring s) {
    SetConsoleTextAttribute(hConsole, GREEN);
    wcout << s << endl;
    SetConsoleTextAttribute(hConsole, WHITE);
}

int InitializeCOM() {
    HRESULT hres;

    // Initialize COM. ------------------------------------------

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x"
            << hex << hres << endl;
        return 1;
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
        exit(EXIT_FAILURE);
    }

    return 0;
}

IWbemLocator* InitializeWMILocator() {
    HRESULT hres;
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
        exit(EXIT_FAILURE);
    }
    return pLoc;
}

IWbemServices* ConnectToWMI(IWbemLocator* pLoc) {
    HRESULT hres;
    IWbemServices* pSvc = NULL;

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
        exit(EXIT_FAILURE);
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
        exit(EXIT_FAILURE);
    }

    return pSvc;
}

// If failed, exits the program
IEnumWbemClassObject* QueryUSBDevices(IWbemServices* pSvc, IWbemLocator* pLoc) {
    HRESULT hres;
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
        exit(EXIT_FAILURE);
    }
    return pEnumerator;
}
