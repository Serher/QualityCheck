//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------







//------------------------------------------------------------------------------
int FindMontajField(Variant &vSheet)
{
    AnsiString sValue;
    int nX = 0;
    for(int nY = 2; nY<=5; nY++)
    {
        do{
            sValue = EX_GetValue(EX_GetCell(vSheet, nY, ++nX));
            if(sValue == "Монтажник")
                return nX;
        }while(nX<25);
        nX = 0;
    }
return 0;
}
//------------------------------------------------------------------------------
int FindAdresField(Variant &vSheet)
{
    AnsiString sValue;
    int nX = 0;
    for(int nY = 2; nY<=5; nY++)
    {
        do{
            sValue = EX_GetValue(EX_GetCell(vSheet, nY, ++nX));
            if(sValue == "Адрес")
                return nX;
        }while(nX<25);
        nX = 0;
    }
return 0;
}
//------------------------------------------------------------------------------
void SetAutoNumber(Variant &vSheet, int nY)
{
    AnsiString sValue = EX_GetValue(EX_GetCell(vSheet, nY-1, 2));
    Variant vCell = EX_GetCell(vSheet, nY, 2);
    if(sValue!="")
        EX_SetValue(vCell, StrToInt(sValue)+1);
    else
        EX_SetValue(vCell, 1);
}
//------------------------------------------------------------------------------
int GetIsCellEmpty(Variant &vCell, TMemo *Memo1)
{
    AnsiString sValue = EX_GetValue(vCell);
    int nColor = EX_GetCellColor(vCell);
    if(sValue != "" || nColor != -4142)
        return false;
    return true;
}
//------------------------------------------------------------------------------
int FindEmptyString(Variant &vSheet, TMemo *Memo1)
{
    Memo1->Lines->Add("Поиск свободной строки");
    int nMin, nY, nMax = 0;
    Variant vCell;
    bool nFound = false;
    nMin = nY = 5;

    do{
        vCell = EX_GetCell(vSheet, nY, 2);
        if(!GetIsCellEmpty(vCell, Memo1))
        {
            nMin = nY;
            if(!nMax)
                nY *=2;
            else
                nY = (nMax - nMin)/2 + nMin;
        }
        else
        {
            nMax = nY;
            nY = (nMax - nMin)/2 + nMin;
        }

        if((nMax-1) == nMin)
        {
            nY = nMax;
            nFound = true;
        }
    }while(!nFound);

    /*
    do{
        nY++;
        vCell = EX_GetCell(vSheet, nY, 2);
        sValue = EX_GetValue(vCell);
        nColor = EX_GetCellColor(vCell);
        Memo1->Lines->Add("sValue: " + sValue + "|| Color: " + IntToStr(nColor));
    }while(sValue != "" || nColor != -4142);
    */
    Memo1->Lines->Add("Найдена строка №"+ IntToStr(nY));
    return nY;
}
//------------------------------------------------------------------------------
int AddSelectedString(Variant &vSourceSheet, int nY, Variant &vTargetSheet, int nAdresField, int nMontajField, TMemo *Memo1, int nTargetY = 0)
{
    //Memo1->Lines->Add("nTargetY " + IntToStr(nTargetY));
    if(!nTargetY)
        nTargetY = FindEmptyString(vTargetSheet, Memo1);
    SetAutoNumber(vTargetSheet, nTargetY);

    EX_SetValue(vTargetSheet, nTargetY, 3, EX_GetValue(EX_GetCell(vSourceSheet, nY, 3)).c_str()); //ФИО
    EX_SetValue(vTargetSheet, nTargetY, 4, EX_GetValue(EX_GetCell(vSourceSheet, nY, 2)).c_str()); //Договор

    if(nAdresField)
        EX_SetValue(vTargetSheet, nTargetY, 5, EX_GetValue(EX_GetCell(vSourceSheet, nY, nAdresField)).c_str()); //Адрес

    if(nMontajField)
        EX_SetValue(vTargetSheet, nTargetY, 6, EX_GetValue(EX_GetCell(vSourceSheet, nY, nMontajField)).c_str()); //Монтажник

    EX_SetValue(vTargetSheet, nTargetY, 7, Now().DateString().c_str()); //Дата проверки

    return nTargetY+1;
}
//------------------------------------------------------------------------------
void CheckSheetForPluses(Variant &vSheet, Variant &vTargetSheet, TMemo *Memo1)
{
    static int nEmptyY;
    int nY = 2;
    int nAdresField = 0;
    int nMontajField = 0;
    Variant vCellWithPlus;
    AnsiString sPlus, sValue = EX_GetValue(EX_GetCell(vSheet, nY, 2));
    while(sValue != "")
    {
        vCellWithPlus = EX_GetCell(vSheet, nY, 1);
        sPlus = EX_GetValue(vCellWithPlus);
        //if(sPlus == "+")
        if(AnsiPos("+", sPlus))
        {
            if(!nAdresField)
                nAdresField = FindAdresField(vSheet);
            if(!nMontajField)
                nMontajField = FindMontajField(vSheet);

            //Memo1->Lines->Add(nMontajField);

            if(EX_GetCellColor(vCellWithPlus) != 6)
            {
                nEmptyY = AddSelectedString(vSheet, nY, vTargetSheet, nAdresField, nMontajField, Memo1, nEmptyY);
                EX_SetCellColor(vCellWithPlus, 6);
            }
        }
        sValue = EX_GetValue(EX_GetCell(vSheet, ++nY, 2));
    }
}
//------------------------------------------------------------------------------





















// Form1
//------------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{
    if(OpenDialog1->Execute())
    {
        Edit1->Text = OpenDialog1->FileName;
    }
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button3Click(TObject *Sender)
{
    if(OpenDialog1->Execute())
    {
        Edit2->Text = OpenDialog1->FileName;
    }
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
    Memo1->Lines->Clear();
    if(Edit1->Text == "" || Edit2->Text == "")
    {
        Memo1->Lines->Add("Один из файлов не выбран!");
        return;
    }

    Button2->Enabled = false;

    ExcelInfo eFile1 = EX_OpenExcelFile(Edit1->Text, false);
    ExcelInfo eFile2 = EX_OpenExcelFile(Edit2->Text, false);

    Variant vSheet, vTargetSheet = EX_GetSheet(eFile2.vBook, 1);
    int nSheetsCount = EX_GetSheetsCount(eFile1.vBook);
    for(int nCount = 1; nCount<=nSheetsCount; nCount++)
    {
        Memo1->Lines->Add("Проверяется лист №" + IntToStr(nCount));
        vSheet = EX_GetSheet(eFile1.vBook, nCount);
        CheckSheetForPluses(vSheet, vTargetSheet, Memo1);
    }

    Memo1->Lines->Add("Проверка закончена");
    EX_SaveFile(eFile1.vBook);
    EX_SaveFile(eFile2.vBook);

    EX_CloseExFile(eFile1.vApp);
    EX_CloseExFile(eFile2.vApp);

    Button2->Enabled = true;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button4Click(TObject *Sender)
{
    if(Edit2->Text == "" || Edit2->Text == "Выберите файл")
    {
        Edit2->Text = "Выберите файл";
        return;
    }

    Memo1->Lines->Clear();
    Memo1->Lines->Add("Ведется подсчет, пожалуйста подождите...");
    Button4->Enabled = false;
    int nYellow, nBlue, nY, nColor;
    nYellow = nBlue = 0;
    nY = 4;

    ExcelInfo eFile2 = EX_OpenExcelFile(Edit2->Text);
    Variant vCell, vSheet = EX_GetSheet(eFile2.vBook, 1);
    do{
        vCell = EX_GetCell(vSheet, ++nY, 3);
        nColor = EX_GetCellColor(vCell);
        if(nColor == 6)
            nYellow++;
        if(nColor == 33)
            nBlue++;

    }while(EX_GetValue(vCell) != "" || (nColor != -4142));
    EX_CloseExFile(eFile2.vApp);

    Memo1->Lines->Add("");
    Memo1->Lines->Add("Желтых строк: " + IntToStr(nYellow));
    Memo1->Lines->Add("Синих строк: " + IntToStr(nBlue));
    Button4->Enabled = true;

    /*
    vCell = EX_GetCell(vSheet, 9, 2);
    Memo1->Lines->Add(EX_GetCellColor(vCell));

    EX_CloseExFile(eFile2.vApp);
    */
}
//---------------------------------------------------------------------------

