#include <iostream>
#include <windows.h>
#include <string>
#include <locale>
#include <cstdlib>


class BaseClass 
{
public:
    std::string title = "    ____  ___    _   ______  ___   _____\n   / __ \\/   |  / | / / __ \\/   | / ___/\n  / /_/ / /| | /  |/ / / / / /| | \\__ \\ \n / ____/ ___ |/ /|  / /_/ / ___ |___/ / \n/_/   /_/  |_/_/ |_/_____/_/  |_|____/  ver 1.0\n";
};

int GetVersionWindows()
{
    // Получаем версию операционной системы
    DWORD majorVersion = 0;
    DWORD minorVersion = 0;
    DWORD buildNumber = 0;
    DWORD platformId = 0;
    std::wstring editionId;

    // Получаем версию операционной системы
    NTSTATUS(WINAPI * RtlGetVersion)(LPOSVERSIONINFOEXW);
    OSVERSIONINFOEXW osInfo;
    ZeroMemory(&osInfo, sizeof(OSVERSIONINFOEXW));
    osInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEXW);

    *(FARPROC*)&RtlGetVersion = GetProcAddress(GetModuleHandleW(L"ntdll"), "RtlGetVersion");
    if (RtlGetVersion != nullptr && RtlGetVersion(&osInfo) == 0)
    {
        majorVersion = osInfo.dwMajorVersion;
        minorVersion = osInfo.dwMinorVersion;
        buildNumber = osInfo.dwBuildNumber;
        platformId = osInfo.dwPlatformId;
    }
    else
    {
        // Не удалось получить информацию о версии операционной системы
        std::cerr << "Error: Could not get OS version information" << std::endl;
        return 1;
    }

    // Получаем редакцию операционной системы
    HKEY hKey;
    DWORD bufferSize = 0;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        if (RegQueryValueExW(hKey, L"EditionID", NULL, NULL, NULL, &bufferSize) == ERROR_SUCCESS)
        {
            wchar_t* buffer = new wchar_t[bufferSize / sizeof(wchar_t)];
            if (RegQueryValueExW(hKey, L"EditionID", NULL, NULL, (LPBYTE)buffer, &bufferSize) == ERROR_SUCCESS)
            {
                editionId.assign(buffer);
            }
            delete[] buffer;
        }
        RegCloseKey(hKey);
    }

    // Definition of system OS version and edition

    std::string editionsWin[5];
    editionsWin[0] = "Core"; // Windows 10 and 11 Home
    editionsWin[1] = "Professional"; // Windows 10 and 11 Professional
    editionsWin[2] = "Enterprise"; // Windows 10 and 11 Enterprise
    editionsWin[3] = "Education"; // Windows 10 and 11 Education
    editionsWin[4] = "CoreSingleLanguage"; // Windows 10 and 11  Home SingleLanguage


    if (platformId == VER_PLATFORM_WIN32_NT && buildNumber < 22000)
    {
        std::string editionIdString(editionId.begin(), editionId.end());
        for (int counter = 0; counter <= 4; counter++)
        {
            if (editionIdString == editionsWin[counter])
            {
                // 100 + counter means that we have Windows 10 with edition ID, which contains in array editionsWin[] by his number
                return 100 + counter;
            }
        }
    }
    else if (platformId == VER_PLATFORM_WIN32_NT && buildNumber >= 22000)
    {
        std::string editionIdString(editionId.begin(), editionId.end());
        for (int counter = 0; counter <= 4; counter++)
        {
            if (editionIdString == editionsWin[counter])
            {
                //110 + counter means that we have Windows 11 with edition ID, which contains in array editionsWin[] by his number
                return 110 + counter;
            }
        }
    }

    return 0;
}

int CommandLineWindows(std::string kmsKey, std::string kmsServer)
{
    std::string firstComm = "slmgr /ipk " + kmsKey;
    std::string secondComm = "slmgr /skms " + kmsServer;
    std::string thirdComm = "slmgr /ato";

    std::system(firstComm.c_str());
    Sleep(3);
    std::system(secondComm.c_str());
    Sleep(10);
    std::system(thirdComm.c_str());
}

int CommandLineOffice(std::string kmsKey, std::string kmsServer, std::string version)
{

    std::string directory = "\"C:\\Program Files\\Microsoft Office\\Office" + version + "\\ospp.vbs\"";
    std::string firstComm = "cscript " + directory + " /inpkey:" + kmsKey;
    std::string secondComm = "cscript " + directory + " /sethst:" + kmsServer;
    std::string thirdComm = "cscript " + directory + " /act";

    std::system(firstComm.c_str());
    Sleep(3);
    std::system(secondComm.c_str());
    Sleep(10);
    std::system(thirdComm.c_str());
}

int ActivationWindows()
{
    // Activation OS Windows 10 and 11

    // Assigning values KMS keys
    std::string kmsKeys[6];
    kmsKeys[0] = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"; // Windows 10 and 11 Home
    kmsKeys[1] = "W269N-WFGWX-YVC9B-4J6C9-T83GX"; // Windows 10 and 11 Professional
    kmsKeys[2] = "NPPR9-FWDCX-D2C8J-H872K-2YT43"; // Windows 10 and 11 Enterprise
    kmsKeys[3] = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"; // Windows 10 and 11 Education
    kmsKeys[4] = "BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"; // Windows 10  Home SingleLanguage
    kmsKeys[5] = "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"; // Windows 11  Home SingleLanguage

    // Array of KMS servers addresses
    std::string kmsServers[8];
    kmsServers[0] = "kms.digiboy.ir";
    kmsServers[1] = "54.223.212.31";
    kmsServers[2] = "kms.cnlic.com";
    kmsServers[3] = "kms.chinancce.com";
    kmsServers[4] = "kms.ddns.net";
    kmsServers[5] = "franklv.ddns.net";
    kmsServers[6] = "k.zpale.com";
    kmsServers[7] = "m.zpale.com";

    int GetVersionWindowsValue = GetVersionWindows();
    int counter;

    if (GetVersionWindowsValue - 100 < 10)
    {
        counter = GetVersionWindowsValue - 100;
        CommandLineWindows(kmsKeys[counter], kmsServers[0]);
    }
    else
    {
        counter = GetVersionWindowsValue - 110;
        CommandLineWindows(kmsKeys[counter], kmsServers[0]);
    }

    return 0;
}

// Activation of Office

int ActivationOffice()
{
    std::cout << "Какая версия Microsoft Office у вас установлена?\n1. Office 2010\n2. Office 2013\n3. Office 2016\n4. Office 2019\n5. Office 2021" << std::endl << "Ввод: ";
    int officeVersion;
    std::string kmsKeys[2];
    int officeLicense;
    std::string version;

    bool flag = false;

    while (flag == false)
    {
        std::cin >> officeVersion;
        system("cls");
        BaseClass header;
        std::cout << header.title;
        if (officeVersion == 1)
        {
            std::cout << "Какая лицензия вам нужна?\n1. Office Professional Plus 2010\n2. Office Standard 2010" << std::endl << "Ввод: ";
            std::cin >> officeLicense;

            // Assigning values KMS keys
            kmsKeys[0] = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB"; // Office Professional Plus 2010
            kmsKeys[1] = "V7QKV-4XVVR-XYV4D-F7DFM-8R6BM"; // Office Standard 2010

            version = "10";

            flag = true;
        }

        else if (officeVersion == 2)
        {
            std::cout << "Какая лицензия вам нужна?\n1. Office Professional Plus 2013\n2. Office Standard 2013" << std::endl << "Ввод: ";
            std::cin >> officeLicense;

            // Assigning values KMS keys
            kmsKeys[0] = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"; // Office Professional Plus 2013
            kmsKeys[1] = "KBKQT-2NMXY-JJWGP-M62JB-92CD4"; // Office Standard 2013

            version = "13";

            flag = true;
        }

        else if (officeVersion == 3)
        {
            std::cout << "Какая лицензия вам нужна?\n1. Office Professional Plus 2016\n2. Office Standard 2016" << std::endl << "Ввод: ";
            std::cin >> officeLicense;

            // Assigning values KMS keys
            kmsKeys[0] = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"; // Office Professional Plus 2016
            kmsKeys[1] = "JNRGM-WHDWX-FJJG3-K47QV-DRTFM"; // Office Standard 2016

            version = "16";

            flag = true;
        }

        else if (officeVersion == 4)
        {
            std::cout << "Какая лицензия вам нужна?\n1. Office Professional Plus 2019\n2. Office Standard 2019" << std::endl << "Ввод: ";
            std::cin >> officeLicense;

            // Assigning values KMS keys
            kmsKeys[0] = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"; // Office Professional Plus 2019
            kmsKeys[1] = "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK"; // Office Standard 2019

            version = "16";

            flag = true;
        }

        else if (officeVersion == 5)
        {
            std::cout << "Какая лицензия вам нужна?\n1. Office LTSC Professional Plus 2021\n2. Office LTSC Standard 2021" << std::endl << "Ввод: ";
            std::cin >> officeLicense;

            // Assigning values KMS keys
            kmsKeys[0] = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"; // Office LTSC Professional Plus 2021
            kmsKeys[1] = "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3"; // Office LTSC Standard 2021

            version = "16";

            flag = true;
        }
        else
        {
            std::cout << "Введены некорректные данные, попробуйте снова:" << std::endl << "Ввод: ";
        }
    }
    

    // Array of KMS servers addresses
    std::string kmsServers[8];
    kmsServers[0] = "kms.digiboy.ir";
    kmsServers[1] = "54.223.212.31";
    kmsServers[2] = "kms.cnlic.com";
    kmsServers[3] = "kms.chinancce.com";
    kmsServers[4] = "kms.ddns.net";
    kmsServers[5] = "franklv.ddns.net";
    kmsServers[6] = "k.zpale.com";
    kmsServers[7] = "m.zpale.com";

    // Clean terminal and add title
    system("cls");
    BaseClass header;
    std::cout << header.title;

    // Activation
    return CommandLineOffice(kmsKeys[officeLicense - 1], kmsServers[0], version);
}

int main()
{
    setlocale(LC_ALL, "Rus");

    BaseClass header;
    std::cout << header.title;

    std::cout << std::endl << "Выберите, что вам нужно активировать:\n1. Microsoft Windows\n2. Microsoft Office" << std::endl << "Ввод: ";
    int input;
    bool flag=false;
    while (flag == false) 
    {
        std::cin >> input;
        if (input == 1)
        {
            system("cls");
            std::cout << header.title;
            return ActivationWindows();
        }
        else if (input == 2)
        {
            system("cls");
            std::cout << header.title;
            return ActivationOffice();
        }
        else
        {
            std::cout << "Введены некорректные данные, попробуйте снова:" << std::endl << "Ввод: ";
        }
    }
}