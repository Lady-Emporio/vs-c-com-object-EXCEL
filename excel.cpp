
#include <tchar.h>
#include <windows.h>
#include <atlbase.h>
#include <iostream>

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\MSO.DLL"
#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB" 
#import "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" \
    rename("DialogBox","_DialogBox") \
    rename("RGB","_RGB") \
    exclude("IFont","IPicture")

using namespace Excel;

int main()
{
    ::CoInitialize(NULL);

    Excel::_ApplicationPtr app("Excel.Application");
    app->Visible[0] = FALSE;
    Excel::_WorkbookPtr book = app->Workbooks->Add();
    Excel::_WorksheetPtr sheet = book->Worksheets->Item[1];

    //Вставляем данные
    sheet->Cells->Item[1, 1] = 3;
    sheet->Cells->Item[1, 2] = 5;
    sheet->Cells->Item[1, 3] = 9;

    int row = 2;
    int col = 1;
    sheet->Cells->Item[row][col] = "Vasa";
    //Создаем диаграмму
    //_ChartPtr  pChart2 = book->Charts->Add();
    //pChart2->ChartWizard((Range*)sheet->Range["A2:C3"], (long)xlLineStacked, 7L, (long)xlRows, 1L, 10L, 5L, "GG");

    //Показываем 
    app->Visible[0] = TRUE;

    return 0;
}