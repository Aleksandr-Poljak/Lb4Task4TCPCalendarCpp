#include "windows.h"
#include <iostream>
#include <atlbase.h>
#include <atlcomcli.h>

int main()
{
    // Инициализация OLE
    HRESULT hr = OleInitialize(NULL);
    if (FAILED(hr))
    {
        std::cout << "Failed to initialize OLE. Code: 0x" << std::hex << hr << "\n";
        return 1;
    }

    // Получаем CLSID COM-сервера
    wchar_t progid[] = L"MSCAL.Календарь";
    CLSID clsid;
    hr = CLSIDFromProgID(progid, &clsid);
    if (FAILED(hr))
    {
        std::cout << "Failed to get CLSID. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }

    // Создаем экземпляр объекта
    CComPtr<IDispatch> pIDispatch;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pIDispatch);
    if (FAILED(hr))
    {
        std::cout << "Failed to create COM instance. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }

    std::cout << "COM instance created successfully.\n";

    // Установка свойства Month
    DISPID dispidMonth;
    OLECHAR* szMonth = (OLECHAR*)L"Month";
    hr = pIDispatch->GetIDsOfNames(IID_NULL, &szMonth, 1, LOCALE_USER_DEFAULT, &dispidMonth);
    if (FAILED(hr))
    {
        std::cout << "Failed to get DISPID for 'Month'. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }

    VARIANTARG varMonth;
    VariantInit(&varMonth);
    varMonth.vt = VT_I4;
    varMonth.lVal = 3;

    DISPPARAMS dispParams = {};
    dispParams.cArgs = 1;
    dispParams.rgvarg = &varMonth;
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    dispParams.cNamedArgs = 1;
    dispParams.rgdispidNamedArgs = &dispidNamed;

    hr = pIDispatch->Invoke(
        dispidMonth,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYPUT,
        &dispParams,
        NULL,
        NULL,
        NULL);

    if (FAILED(hr))
    {
        std::cout << "Failed to set 'Month'. Code: 0x" << std::hex << hr << "\n";
    }
    else
    {
        std::cout << "'Month' set to 3 successfully.\n";
    }

    VariantClear(&varMonth);

    DISPID dispidYear;
    OLECHAR* szYear = (OLECHAR*)L"Year";
    hr = pIDispatch->GetIDsOfNames(IID_NULL, &szYear, 1, LOCALE_USER_DEFAULT, &dispidYear);
    if (FAILED(hr))
    {
        std::cout << "Failed to get DISPID for 'Year'. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }

    VARIANTARG varYear;
    VariantInit(&varYear);
    varYear.vt = VT_I4;
    varYear.lVal = 2026;

    dispParams.rgvarg = &varYear;

    hr = pIDispatch->Invoke(
        dispidYear,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYPUT,
        &dispParams,
        NULL,
        NULL,
        NULL);

    if (FAILED(hr))
    {
        std::cout << "Failed to set 'Year'. Code: 0x" << std::hex << hr << "\n";
    }
    else
    {
        std::cout << "'Year' set to 2026 successfully.\n";
    }
    VariantClear(&varYear);

    // Получение значения Day
    DISPID dispidDay;
    OLECHAR* nameDay = (OLECHAR*)L"Day";
    hr = pIDispatch->GetIDsOfNames(IID_NULL, &nameDay, 1, LOCALE_USER_DEFAULT, &dispidDay);
    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    if (FAILED(hr))
    {
        std::cout << "Failed to get DISPID for 'Day'. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }

    VARIANT resultDay;
    VariantInit(&resultDay);

    hr = pIDispatch->Invoke(dispidDay, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &resultDay, NULL, NULL);
    if (SUCCEEDED(hr))
    {
        std::cout << "Day: " << resultDay.lVal << "\n";
    }
    else
    {
        std::cout << "Failed to get 'Day'. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }
    VariantClear(&resultDay);

    VARIANT resultMonth;
    VariantInit(&resultMonth);
    hr = pIDispatch->Invoke(dispidMonth, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &resultMonth, NULL, NULL);
    if (SUCCEEDED(hr))
    {
        std::cout << "Month: " << resultMonth.lVal << "\n";
    }
    else
    {
        std::cout << "Failed to get 'Month'. Code: 0x" << std::hex << hr << "\n";
    }
    VariantClear(&resultMonth);



    VARIANT resultYear;
    VariantInit(&resultYear);
    hr = pIDispatch->Invoke(dispidYear, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &resultYear, NULL, NULL);
    if (SUCCEEDED(hr))
    {
        std::cout << "Year: " << resultYear.lVal << "\n";
    }
    else
    {
        std::cout << "Failed to get 'Year'. Code: 0x" << std::hex << hr << "\n";
        OleUninitialize();
        return 1;
    }
    VariantClear(&resultYear);

    // Завершение работы
    pIDispatch.Release();
    OleUninitialize();
    return 0;
}
