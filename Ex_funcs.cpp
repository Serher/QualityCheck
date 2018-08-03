#include <ComObj.hpp>

//---------------------------------------------------------------------------
//----Variant Functions------------------------------------------------------
//---------------------------------------------------------------------------
class ExcelInfo
{
    public:
    Variant vApp;
    Variant vBook;
};
//---------------------------------------------------------------------------
Variant EX_InitApp()
{
    Variant vApp;
    try
    {
        vApp=CreateOleObject("Excel.Application");
    }
    catch(...)
    {
        MessageBox(0, "Ошибка при открытии сервера Excel",
            "Ошибка", MB_OK);
        return 1;
    }
return vApp;
}
//---------------------------------------------------------------------------
void EX_ShowExcel(Variant &vApp)
{
    vApp.OlePropertySet("Visible",true);
}
//---------------------------------------------------------------------------
Variant EX_GetWorkBook(Variant &vApp, AnsiString &vasPath, int nReadOnly = true)
{
    Variant vBook, vBooks = vApp.OlePropertyGet("Workbooks");

    //Adres correction
    AnsiString sSub = "\\";
    sSub.Insert("\\", 1);
    if(!AnsiPos(sSub, vasPath))
    {
        int nLen = vasPath.Length();
        while(nLen>1)
        {
            if(vasPath.SubString(nLen, 1) == '\\')
                vasPath.Insert('\\', nLen);
            nLen--;
        }
    }
    //Adres correction end

    vBooks.OleProcedure("Open", vasPath.c_str(), 0, nReadOnly);
    vBook = vBooks.OlePropertyGet("Item",1);
return vBook;
}
//---------------------------------------------------------------------------
Variant EX_GetSheet(Variant &vBook, int nNumber)
{
    Variant vSheets = vBook.OlePropertyGet("Worksheets");
return vSheets.OlePropertyGet("Item", nNumber);
}
//---------------------------------------------------------------------------
int EX_GetSheetsCount(Variant &vBook)
{
    Variant vSheets = vBook.OlePropertyGet("Worksheets");
return vSheets.OlePropertyGet("Count");
}
//---------------------------------------------------------------------------
Variant EX_GetCell(Variant &vSheet, int nY, int nX)
{
	return vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nY,nX);
}
//---------------------------------------------------------------------------
AnsiString EX_GetValue(Variant &vCell)
{
    return VarToStr(vCell);
}
//---------------------------------------------------------------------------
void EX_SetValue(Variant &vSheet, int nY, int nX, char *sValue)
{
    Variant vCell = EX_GetCell(vSheet, nY, nX);
    vCell.OlePropertySet("Value", sValue);
}
//---------------------------------------------------------------------------
void EX_SetValue(Variant &vSheet, int nY, int nX, int nValue)
{
    Variant vCell = EX_GetCell(vSheet, nY, nX);
    vCell.OlePropertySet("Value", nValue);
}
//---------------------------------------------------------------------------
void EX_SetValue(Variant &vCell, char *sValue)
{
    vCell.OlePropertySet("Value", sValue);
}
//---------------------------------------------------------------------------
void EX_SetValue(Variant &vCell, int nValue)
{
    vCell.OlePropertySet("Value", nValue);
}
//---------------------------------------------------------------------------
void EX_SaveFile(Variant &vBook)
{
    vBook.OleProcedure("Save");
}
//---------------------------------------------------------------------------
void EX_CloseExFile(Variant &vApp)
{
    vApp.OleProcedure("Quit");
}
//---------------------------------------------------------------------------
int EX_GetCellColor(Variant &vCell)
{
    return vCell.OlePropertyGet("Interior").OlePropertyGet("ColorIndex");
}
//---------------------------------------------------------------------------
void EX_SetCellColor(Variant &vCell, int nColor) //0-Black, 1-white, ?? 3-R, 4-G, 5-B, 6-Yellow, etc.
{
    vCell.OlePropertyGet("Interior").OlePropertySet("ColorIndex", nColor);
}
//---------------------------------------------------------------------------
ExcelInfo EX_OpenExcelFile(AnsiString &sPath, int nReadOnly = true)
{
    ExcelInfo eResult;
    eResult.vApp = EX_InitApp();
    eResult.vBook = EX_GetWorkBook(eResult.vApp, sPath, nReadOnly);
return eResult;
}
//---------------------------------------------------------------------------
//--End of Variant functions-------------------------------------------------
//---------------------------------------------------------------------------