<cfscript>
itemAgreementStyle = StructNew();
itemAgreementStyle.font = "Arial";
itemAgreementStyle.fontsize = "20";
itemAgreementStyle.alignment = "center";
itemAgreementStyle.bold = "true";
itemAgreementStyle.italic = "true";
itemAgreementStyle.textwrap = "true";
itemAgreementStyle2.font = "Arial";
itemAgreementStyle2.fontsize = "16";
itemAgreementStyle2.alignment = "center";
itemAgreementStyle2.bold = "true";
itemAgreementStyle2.italic = "true";
itemAgreementStyle2.textwrap = "true";
theFile = "#DateFormat(now(), 'DD-MM-YYYY')#.xlsx";
// Create a new spreadsheet
spreadsheet = spreadsheetNew("Gemstones_ Pearl Quote Sheet", true);
spreadsheetCreateSheet(spreadsheet, 'Diamond Quote Sheet');
spreadSheetSetActiveSheet(spreadsheet, 'Diamond Quote Sheet');
imagePath = expandPath("images/Costco_Logo.png"); 
spreadsheetAddImage(spreadsheet, imagePath, "1, 1, 2, 2");
spreadsheetSetColumnWidth(spreadsheet, 1, 33);
spreadsheetMergeCells(spreadsheet, 1, 1, 6, 10); 
// Set "ITEM AGREEMENT" in the first row
spreadsheetSetCellValue(spreadsheet, "ITEM AGREEMENT", 1, 6);
spreadsheetFormatCell(spreadsheet, itemAgreementStyle, 1, 6);
spreadsheetMergeCells(spreadsheet, 2, 2, 6, 10); 
// Set "JEWELERY QUOTE FORM" in the second row
spreadsheetSetCellValue(spreadsheet, "JEWELERY QUOTE FORM", 2, 6);
spreadsheetFormatCell(spreadsheet, itemAgreementStyle2, 2, 6);
// Set "SUPPLIER INFORMATION" in the third row 
spreadsheetMergeCells(spreadsheet, 3, 3, 1, 6);
spreadsheetSetCellValue(spreadsheet, "SUPPLIER INFORMATION", 3, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", fgcolor="yellow", color="black", alignment="center"}, 3, 1);
// Apply borders specifically to the third row
for (col = 1; col <= 6; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin'}, 3, col);
}
//set vendor text information in fifth row
spreadsheetSetCellValue(spreadsheet, 'VENDOR ##',  5, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black", alignment="right"},5,1)
//set borderbottom line for the vendor row
spreadsheetMergeCells(spreadsheet, 5, 5, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  5, 6);
for (col = 1; col <= 6; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 5, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 5, col);
    }
}

spreadsheetSetCellValue(spreadsheet, 'QUOTE PROVIDED BY (NAME) :',  5, 9);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black", alignment="right"},5,9);
spreadsheetMergeCells(spreadsheet, 5, 5, 10, 14);
spreadsheetSetCellValue(spreadsheet, '',  5, 16);


for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 5, col);
}
//set vendor/company name in the 6th row:
spreadsheetSetCellValue(spreadsheet, 'VENDOR / COMPANY NAME: ',  6, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},6,1);
//set borderbottom line for the vendor row
spreadsheetMergeCells(spreadsheet, 6, 6, 2, 6);

spreadsheetSetCellValue(spreadsheet, '',  6, 6);
for (col = 1; col <= 6; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 6, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 6, col);
    }
}
spreadsheetSetCellValue(spreadsheet, 'POSITION:',  6, 9);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black", alignment="right"},6,9);
spreadsheetMergeCells(spreadsheet, 6, 6, 10, 14);
spreadsheetSetCellValue(spreadsheet, '',  6, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 6, col);
}
//set ADDRESS name in the 7th row:
spreadsheetSetCellValue(spreadsheet, 'ADDRESS:',  7, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black",alignment="right"},7,1);
//set borderbottom line for the email row
spreadsheetMergeCells(spreadsheet, 7, 7, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  7, 6);
for (col = 1; col <= 6; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 7, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 7, col);
    }
}
spreadsheetSetCellValue(spreadsheet, 'EMAIL:',  7, 9);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black", alignment="right"},7,9);
spreadsheetMergeCells(spreadsheet, 7, 7, 10, 14);
spreadsheetSetCellValue(spreadsheet, '',  7, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 7, col);
}

//set CITY/STATE/ZIP name in the 8th row:
spreadsheetSetCellValue(spreadsheet, 'CITY / STATE / ZIP:',  8, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},8,1);
//set borderbottom line for the 8th row
spreadsheetMergeCells(spreadsheet, 8, 8, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  8, 6);
for (col = 1; col <= 6; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 8, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 8, col);
    }
}
spreadsheetSetCellValue(spreadsheet, 'QUOTE IS VALID FOR WHICH COUNTRIES:', 8, 8);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black"}, 8, 8);
spreadsheetSetColumnWidth(spreadsheet, 8, 24);
spreadsheetMergeCells(spreadsheet, 8, 8, 10, 14);
spreadsheetSetCellValue(spreadsheet, '',  8, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 8, col);
}
//set telephone name in the 9th row:
spreadsheetSetCellValue(spreadsheet, 'TELEPHONE ##:',  9, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},9,1);
//set borderbottom line for the 9th row
spreadsheetMergeCells(spreadsheet, 9, 9, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  9, 6);
for (col = 1; col <= 6; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 9, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 9, col);
    }
}
// Set "ITEM INFORMATION" in the 11th row and apply format
spreadsheetMergeCells(spreadsheet, 11, 11, 1, 6);
spreadsheetSetCellValue(spreadsheet, "ITEM INFORMATION", 11, 1);
    spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="yellow", color="black", alignment="center"}, 11, 1);
for (col = 1; col <= 6; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin'}, 11, col);
}
//set PURCHASE ORDER INFORMATION in the 11th row
spreadsheetMergeCells(spreadsheet, 11, 11, 9, 14);
spreadsheetSetCellValue(spreadsheet, "PURCHASE ORDER INFORMATION", 11, 9);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="yellow", color="black", alignment="center", bgcolor="yellow"}, 11, 9);
for (col = 9; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin'}, 11, col);
    
}

//set dimensions  in the 12th row
spreadsheetSetCellValue(spreadsheet, 'Dimensions :',  12, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="right"},12,1);
spreadsheetSetCellValue(spreadsheet, 'Height:', 12, 7);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black",alignment="center"}, 12, 7);
spreadsheetSetColumnWidth(spreadsheet, 7, 15);
spreadsheetMergeCells(spreadsheet, 12, 12, 9, 11);
spreadsheetSetCellValue(spreadsheet, 'PURCHASE ORDER NUMBER', 12, 9);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 12, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin',bgcolor="none"}, 12, col);
}
spreadsheetMergeCells(spreadsheet, 12, 12, 12, 14);
spreadsheetSetCellValue(spreadsheet, 'QUANTITY & SHIP DATE', 12, 12);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 12, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',bgcolor="none"}, 12, col);
}
//set l,w,h,cube in the 13th row
spreadsheetSetCellValue(spreadsheet, 'L :',  13, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="right"},13,1);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 2);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 2);
spreadsheetSetCellValue(spreadsheet, 'W:',  13, 3);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="center"},13,3);
spreadsheetSetColumnWidth(spreadsheet, 3, 5);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 4);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 4);
spreadsheetSetCellValue(spreadsheet, 'H :',  13, 5);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="center"},13,5);
spreadsheetSetColumnWidth(spreadsheet, 5, 5);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 6);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 6);
spreadsheetSetCellValue(spreadsheet, 'H :',  13, 5);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="right"},13,5);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 6);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 6);
spreadsheetSetCellValue(spreadsheet, 'Cube:', 13, 7);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black",alignment="center"}, 13, 7);
spreadsheetSetCellValue(spreadsheet, "0.00", 13, 8);
formatStruct=structNew();
formatStruct.alignment = "center";
formatStruct.fontsize = "12";
formatStruct.dataformat = "0.00";
spreadsheetFormatCellRange(spreadsheet,formatStruct, 13, 8, 13, 8);
spreadsheetSetCellValue(spreadsheet, '', 13, 9);
spreadsheetFormatCell(spreadsheet,{color="black"},13, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none"}, 13, col);
}
spreadsheetMergeCells(spreadsheet, 13, 13, 9, 14);
spreadsheetSetCellValue(spreadsheet, '', 13, 12);
spreadsheetFormatCell(spreadsheet, {color="black"},13, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none"}, 13, col);
}

//set borderlines and purchase order number and ship date in the 14th row
spreadsheetMergeCells(spreadsheet, 14, 14, 9, 11);
spreadsheetSetCellValue(spreadsheet, 'PURCHASE ORDER NUMBER', 14, 9);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 14, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', rightborder='thin',leftborder='thin',bgcolor="none"}, 14, col);
}
spreadsheetMergeCells(spreadsheet, 14, 14, 12, 14);
spreadsheetSetCellValue(spreadsheet, 'QUANTITY & SHIP DATE', 14, 12);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 14, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thin',bgcolor="none"}, 14, col);
}

//set costco depost information in the 15th row
spreadsheetSetCellValue(spreadsheet, 'Costco Depot(889 / 894 / BOTH):', 15, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true}, 15, 1);
//set borderbottom line for the costco depost information(15th) row
spreadsheetMergeCells(spreadsheet, 15, 15, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  15, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 15, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 15, col);
    }
}
//SET EX: values in the 15th row
spreadsheetMergeCells(spreadsheet, 15, 15, 9, 11);
spreadsheetSetCellValue(spreadsheet, 'EX:8950101123', 15, 9);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 15, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', rightborder='thin',leftborder='thin',bgcolor="none"}, 15, col);
}
spreadsheetMergeCells(spreadsheet, 15, 15, 12, 14);
spreadsheetSetCellValue(spreadsheet, '11/1/31', 15, 12);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 15, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thin',bgcolor="none",alignment="center"}, 15, col);
}

//set reorder in the 16th row
spreadsheetSetCellValue(spreadsheet, 'Reorder(Y/N):', 16, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 16, 1);
spreadsheetMergeCells(spreadsheet, 16, 16, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  16, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 16, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 16, col);
    }
}
//set border line in the 16th row empty contents
spreadsheetMergeCells(spreadsheet, 16, 16, 9, 11)
spreadsheetSetCellValue(spreadsheet, '', 16, 9);
spreadsheetFormatCell(spreadsheet,{color="black"},16, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin"}, 16, col);
}
spreadsheetMergeCells(spreadsheet, 16, 16, 12, 14);
spreadsheetSetCellValue(spreadsheet, '', 16, 12);
spreadsheetFormatCell(spreadsheet, {color="black"},16, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 16, col);
}


//set new item in the 17th row
spreadsheetSetCellValue(spreadsheet, 'New Item (Y/N):', 17, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 17, 1);
spreadsheetMergeCells(spreadsheet, 17, 17, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  17, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 17, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 17, col);
    }
}
//set border line in the 17th row empty contents
spreadsheetMergeCells(spreadsheet, 17, 17, 9, 11)
spreadsheetSetCellValue(spreadsheet, '', 17, 9);
spreadsheetFormatCell(spreadsheet,{color="black"},17, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin"}, 17, col);
}
spreadsheetMergeCells(spreadsheet, 17, 17, 12, 14);
spreadsheetSetCellValue(spreadsheet, '', 17, 12);
spreadsheetFormatCell(spreadsheet, {color="black"},17, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 17, col);
}
//set item description in the 18th row
spreadsheetSetCellValue(spreadsheet, 'Item Description:', 18, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 18, 1);
spreadsheetMergeCells(spreadsheet, 18, 18, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  18, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 18, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 18, col);
    }
}
//set border line in the 18th row empty contents
spreadsheetMergeCells(spreadsheet, 18, 18, 9, 11)
spreadsheetSetCellValue(spreadsheet, '', 18, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin"}, 18, col);
}
spreadsheetMergeCells(spreadsheet, 18, 18, 12, 14);
spreadsheetSetCellValue(spreadsheet, '', 18, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin',color="black"}, 18, col);
}

//set vendor style # in the 19th row
spreadsheetSetCellValue(spreadsheet, 'Vendor Style ##:', 19, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 19, 1);
spreadsheetMergeCells(spreadsheet, 19, 19, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  19, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 19, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 19, col);
    }
}
//set border line in the 19th row empty contents
spreadsheetMergeCells(spreadsheet, 19, 19, 9, 11)
spreadsheetSetCellValue(spreadsheet, '', 19, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin",color="black"}, 19, col);
}
spreadsheetMergeCells(spreadsheet, 19, 19, 12, 14);
spreadsheetSetCellValue(spreadsheet, '', 19, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin',color="black"}, 19, col);
}

//set minimum cwt in the 20th row
spreadsheetSetCellValue(spreadsheet, 'Minimum CWT:', 20, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 20, 1);
spreadsheetMergeCells(spreadsheet, 20, 20, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  20, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 20, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 20, col);
    }
}
//set border line in the 20th row empty contents
spreadsheetMergeCells(spreadsheet, 20, 20, 9, 11)
spreadsheetSetCellValue(spreadsheet, '', 20, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin",color="black"}, 20, col);
}
spreadsheetMergeCells(spreadsheet, 20, 20, 12, 14);
spreadsheetSetCellValue(spreadsheet, '', 20, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin',color="black"}, 20, col);
}

//set minimum cwt in the 21st row
spreadsheetSetCellValue(spreadsheet, 'Minimum Center CWT:', 21, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 21, 1);
spreadsheetMergeCells(spreadsheet, 21, 21, 2, 7);
spreadsheetSetCellValue(spreadsheet, '',  21, 6);
for (col = 1; col <= 7; col++) {
    if (col == 1) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='none'}, 21, col);
    } else {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 21, col);
    }
}
//set costco item numbers in the 21st row
spreadsheetMergeCells(spreadsheet, 21, 21, 9, 14);
spreadsheetSetCellValue(spreadsheet, "COSTCO ITEM NUMBERS(S)", 21, 9);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="yellow", color="black", alignment="center", bgcolor="yellow"}, 21, 9);
for (col = 9; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin'}, 21, col);
    
}
//set item/features information in the 22nd row
spreadsheetSetCellValue(spreadsheet, 'Item Features/Specs:', 22, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 22, 1);
spreadsheetMergeCells(spreadsheet, 22, 22, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 22, col);
}
spreadsheetMergeCells(spreadsheet, 22, 22, 6, 7);
for(col=6;col<=7;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 22, col);
}
//set 22nd row(secondset)
spreadsheetMergeCells(spreadsheet, 22, 22, 9, 11);
spreadsheetSetCellValue(spreadsheet, '##1', 22, 9);
spreadsheetFormatCell(spreadsheet,{color="black"},22, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",alignment="center"}, 22, col);
}
spreadsheetMergeCells(spreadsheet, 22, 22, 12, 14);
spreadsheetSetCellValue(spreadsheet, '##2', 22, 12);
spreadsheetFormatCell(spreadsheet, {color="black"},22, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",alignment="center"}, 22, col);
}
//set empty contents in the 23rd row
spreadsheetSetCellValue(spreadsheet, '', 23, 1);
spreadsheetMergeCells(spreadsheet, 23, 23, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 23, col);
}
spreadsheetMergeCells(spreadsheet, 23, 23, 6, 7);
for(col=6;col<=7;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 23, col);
}

//set 23rd row (secondSet)
spreadsheetMergeCells(spreadsheet, 23, 23, 9, 11);
spreadsheetSetCellValue(spreadsheet, '##3', 23, 9);
spreadsheetFormatCell(spreadsheet,{color="black"},23, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",alignment="center"}, 23, col);
}
spreadsheetMergeCells(spreadsheet, 23, 23, 12, 14);
spreadsheetSetCellValue(spreadsheet, '##4', 23, 12);
spreadsheetFormatCell(spreadsheet, {color="black"},23, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",alignment="center"}, 23, col);
}
//set empty contents in the 24th row
spreadsheetSetCellValue(spreadsheet, '', 24, 1);
spreadsheetMergeCells(spreadsheet, 24, 24, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 24, col);
}
spreadsheetMergeCells(spreadsheet, 24, 24, 6, 7);
for(col=6;col<=7;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 24, col);
}
//set 24th row(secondset)
spreadsheetMergeCells(spreadsheet, 24, 24, 9, 11);
spreadsheetSetCellValue(spreadsheet, '##5', 24, 9);
spreadsheetFormatCell(spreadsheet,{color="black"},24, 9);
for(col=9;col<=11;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',bgcolor="none",alignment="center"}, 24, col);
}
spreadsheetMergeCells(spreadsheet, 24, 24, 12, 14);
spreadsheetSetCellValue(spreadsheet, '##6', 24, 12);
spreadsheetFormatCell(spreadsheet, {color="black"},24, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bottomborder='thin',bgcolor="none",alignment="center"}, 24, col);
}
//set empty contents in the 25th row
spreadsheetSetCellValue(spreadsheet, '', 25, 1);
spreadsheetFormatCell(spreadsheet, {}, 25, 1);
spreadsheetMergeCells(spreadsheet, 25, 25, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 25, col);
}
spreadsheetMergeCells(spreadsheet, 25, 25, 6, 7);
for(col=6;col<=7;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick'}, 25, col);
}
//set image information in the 26th row(secondset)
spreadsheetSetCellValue(spreadsheet, 'IMAGE:', 26, 9);
spreadsheetFormatCell(spreadsheet, {alignment='left'}, 26, 9);

//set item cost details in the 27th row 
spreadsheetMergeCells(spreadsheet, 27, 27, 1, 6);
spreadsheetSetCellValue(spreadsheet, "ITEM COST DETAILS", 27, 1);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="yellow", color="black", alignment="center"}, 27, 1);
for (col = 1; col <= 6; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thick',rightborder='thick'}, 27, col);
}
//SET empty image set in the 27th-35th row
spreadsheetMergeCells(spreadsheet, 27, 35, 9, 14);
spreadsheetSetCellValue(spreadsheet, '', 27, 9);
spreadsheetFormatCell(spreadsheet, {bgcolor='none'}, 27, 9);
for (col = 9; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder: 'thin', topborder: 'thin', rightborder: 'thin', leftborder: 'thin', bgcolor='none'}, 27, col);
    for (row = 28; row <= 34; row++) {
        spreadsheetFormatCell(spreadsheet, {topborder: 'thin', rightborder: 'thin', leftborder: 'thin', bgcolor='none'}, row, col);
    }
    spreadsheetFormatCell(spreadsheet, {bottomborder: 'thin', topborder: 'thin', rightborder: 'thin', leftborder: 'thin', bgcolor='none'}, 35, col);
}

//set quote data details in the 28th row
spreadsheetSetCellValue(spreadsheet, "QUOTE DATE:", 28, 1);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black", alignment="right"}, 28, 1);
spreadsheetSetCellValue(spreadsheet, '', 28, 2);
spreadsheetMergeCells(spreadsheet, 28, 28, 2, 6);
for(col=2;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='thick'}, 28, col);
}
//set usmca data details in the 29th row
spreadsheetSetCellValue(spreadsheet, "USMCA APPLICABLE (Y/N):", 29, 1);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black", alignment="right"}, 29, 1);
spreadsheetSetCellValue(spreadsheet, '', 29, 2);
spreadsheetMergeCells(spreadsheet, 29, 29, 2, 6);
for(col=2;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='thick'}, 29, col);
}
//SET EMPTY CONTENT IN THE 30th row
spreadsheetSetCellValue(spreadsheet, "", 30, 1);
spreadsheetFormatCell(spreadsheet, {}, 30, 1);
spreadsheetSetCellValue(spreadsheet, '', 30, 2);
spreadsheetMergeCells(spreadsheet, 30, 30, 2, 6);
for(col=2;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='thick'}, 30, col);
}
//SET price at  CONTENT IN THE 31sth row
spreadsheetSetCellValue(spreadsheet, "PRICED AT:", 31, 1);
spreadsheetFormatCell(spreadsheet, {bold=true,rightborder='thin',alignment='right'}, 31, 1);
spreadsheetSetCellValue(spreadsheet, 'Gold:', 31, 2);
spreadsheetFormatCell(spreadsheet, {alignment='right',bottomborder='thin',topborder='thin',rightborder='thin'}, 31, 2);
spreadsheetSetColumnWidth(spreadsheet, 2, 10);
spreadsheetMergeCells(spreadsheet, 31, 31, 3, 6);
spreadsheetSetCellValue(spreadsheet, '', 31, 3);
    for(col=3;col<=6;col++){
        spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='thick'}, 31, col);
    }

//set platinum content in the 32nd row
spreadsheetMergeCells(spreadsheet, 32, 32, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Platinum:', 32, 1);
spreadsheetFormatCell(spreadsheet, {alignment="right"}, 32, 1);
spreadsheetMergeCells(spreadsheet, 32, 32, 3, 6)
spreadsheetSetCellValue(spreadsheet, '', 32, 3);
for(col=3;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thick',bottomborder='thin',leftborder='thin'}, 32, col);
}
//set minimum cwt in the 33rd row
spreadsheetMergeCells(spreadsheet, 33, 33, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Minimum CWT:', 33, 1);
spreadsheetFormatCell(spreadsheet, {alignment="right"}, 33, 1);
spreadsheetMergeCells(spreadsheet, 33, 33, 3, 6)
spreadsheetSetCellValue(spreadsheet, '', 33, 3);
for(col=3;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thick',bottomborder='thin',leftborder='thin'}, 33, col);
}
//SET EMPTY contents in the 34th row
spreadsheetMergeCells(spreadsheet, 34, 34, 1, 2);
spreadsheetSetCellValue(spreadsheet, '', 34, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder="thick"}, 34, col);
}

spreadsheetMergeCells(spreadsheet, 34, 34, 3, 6)
spreadsheetSetCellValue(spreadsheet, '', 34, 3);
for(col=3;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thick',bottomborder='thick',leftborder='thin'}, 34, col);
}
//set mounting information in the 35th row
spreadsheetMergeCells(spreadsheet, 35, 35, 1, 6);
spreadsheetSetCellValue(spreadsheet, 'MOUNTING:', 35, 1);
spreadsheetformatcell(spreadsheet,{alignment="left",bold=true,color="black"},35,1);
for(col=1;col<=6;col++){
    spreadsheetformatcell(spreadsheet,{rightborder='thick'},35,col);
}
//set finished dwt in 36th row
spreadsheetMergeCells(spreadsheet, 36, 36, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Finished DWT', 36, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 36, 1);
spreadsheetMergeCells(spreadsheet, 36, 36, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 36, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 36, col);
}
//set casting charge in the 37th row
spreadsheetMergeCells(spreadsheet, 37, 37, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Casting Charge', 37, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 37, 1);
spreadsheetMergeCells(spreadsheet, 37, 37, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 37, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 37, col);
}
//set gold breakdown title in the 37th row(secondSet)
spreadsheetMergeCells(spreadsheet, 37, 37, 8, 9);
spreadsheetSetCellValue(spreadsheet, 'Gold Breakdown', 37, 8);
formatGoldValue=structNew();
formatGoldValue.color='black';
formatGoldValue.bold=true;
formatGoldValue.alignment='center';
formatGoldValue.fontsize='14';
formatGoldValue.bgcolor='yellow';
formatGoldValue.fgcolor='yellow';
formatGoldValue.topborder='thick';
formatGoldValue.bottomborder='thin';
formatGoldValue.rightborder='thick';
formatGoldValue.leftborder='thick';
for(col=8;col<=9;col++){
    spreadsheetFormatCell(spreadsheet, formatGoldValue, 37, col);
}
//set finding/chain in the 38th row
spreadsheetMergeCells(spreadsheet, 38, 38, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Finding / Chain', 38, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 38, 1);
spreadsheetMergeCells(spreadsheet, 38, 38, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 38, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 38, col);
}
//set gram informaion in the 38th row(secondset)
spreadsheetSetCellValue(spreadsheet, 'Gram:', 38, 8);
spreadsheetFormatCell(spreadsheet, {alignment='left',bgcolor='yellow',fgcolor='yellow',fontsize='13',leftborder='thick',rightborder='thin',bottomborder='thin'}, 38, 8);
spreadsheetSetCellValue(spreadsheet, '', 38, 9);
spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 38, 9);
//set pcs on Casting in the 39th row
spreadsheetMergeCells(spreadsheet, 39, 39, 1, 2);
spreadsheetSetCellValue(spreadsheet, '## Pcs on Casting', 39, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 39, 1);
spreadsheetMergeCells(spreadsheet, 39, 39, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 39, 4);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 39, 1);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 39, col);
}
//set labour informaion in the 39th row(secondset)
spreadsheetSetCellValue(spreadsheet, 'Labour:', 39, 8);
spreadsheetFormatCell(spreadsheet, {alignment='left',bgcolor='yellow',fgcolor='yellow',fontsize='13',leftborder='thick',rightborder='thin',bottomborder='thin'}, 39, 8);
spreadsheetSetCellValue(spreadsheet, '', 39, 9);
spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 39, 9);
//set head size/shape on Casting in the 40th row
spreadsheetMergeCells(spreadsheet, 40, 40, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Head Size / Shape', 40, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 40, 1);
spreadsheetMergeCells(spreadsheet, 40, 40, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 40, 4);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 40, 1);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 40, col);
}
//set $ per gram informaion in the 40th row(secondset)
spreadsheetSetCellValue(spreadsheet, '$ Per Gram', 40, 8);
spreadsheetFormatCell(spreadsheet, {alignment='left',bgcolor='yellow',fgcolor='yellow',fontsize='13',leftborder='thick',rightborder='thin',bottomborder='thick'}, 40, 8);
spreadsheetSetCellValue(spreadsheet, '$0.00', 40, 9);
spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick',alignment='center',bottomborder='thick'}, 40, 9);
//set total mounting in the 41st row
spreadsheetMergeCells(spreadsheet, 41, 41, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Total Mounting', 41, 1);
for(col=1;col<=2;col++){
    spreadsheetformatcell(spreadsheet,{alignment="right",bold=true,color="black"},41,col);
}
//set value $0.00 in the 41st row
spreadsheetMergeCells(spreadsheet, 41, 41, 4, 6);
spreadsheetSetCellValue(spreadsheet, "$0.00", 41, 4);
formatValue=structNew();
formatValue.alignment = "center";
formatValue.fontsize = "12";
formatValue.dataformat = "$0.00";
formatValue.bold="true";
formatValue.color="black";
formatValue.rightborder="thick";
for(col=4;col<=6;col++){
    spreadsheetFormatCellRange(spreadsheet,formatValue, 41, 4, 41, 6);
}
//set empty bottom border line in the 42nd row
spreadsheetMergeCells(spreadsheet, 42, 42, 1, 6)
spreadsheetSetCellValue(spreadsheet, '', 42, 1);
for(col=1;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick',rightborder='thick'}, 42, col);
}
//set labor cost details in the 43rd row(1st set)
spreadsheetMergeCells(spreadsheet, 43, 43, 1, 6);
spreadsheetSetCellValue(spreadsheet, 'LABOR COSTS :', 43, 1);
for(col=1;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='left',fontsize='12',bold='true',color='black',rightborder='thick'}, 43, col)
}
//set labor cost details in the 43rd row(2nd set)
spreadsheetMergeCells(spreadsheet, 43, 43, 8, 14);
spreadsheetSetCellValue(spreadsheet, 'LABOR COST DETAILS ', 43, 8);
formatLaborValue=structNew();
formatLaborValue.alignment='center';
formatLaborValue.bgcolor='yellow';
formatLaborValue.fgcolor='yellow';
formatLaborValue.topborder='thick';
formatLaborValue.bottomborder='thick';
formatLaborValue.rightborder='thick';
formatLaborValue.leftborder='thick';
formatLaborValue.fontsize='13';
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, formatLaborValue, 43, col);
}
//set cost to assemble value for the 44th row(cost to assemble)
spreadsheetMergeCells(spreadsheet, 44, 44, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Cost To Assemble', 44, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 44, 1);
spreadsheetMergeCells(spreadsheet, 44, 44, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 44, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 44, col);
}
//set empty cell for the 44th row(second set)
spreadsheetMergeCells(spreadsheet, 44, 44, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 44, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='thick',rightborder='thick'}, 44, col);
}
//set what needs to be assembled value for the 45th row
spreadsheetMergeCells(spreadsheet, 45, 45, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'What Needs to be Assembled', 45, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 45, 1);
spreadsheetMergeCells(spreadsheet, 45, 45, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 45, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 45, col);
}
//set empty cell for the 45th row(2nd set)
spreadsheetMergeCells(spreadsheet, 45, 45, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 45, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='thick',rightborder='thick'}, 45, col);
}
//set polish and finish value for the 46th row
spreadsheetMergeCells(spreadsheet, 46, 46, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Polish & Finish', 46, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 46, 1);
spreadsheetMergeCells(spreadsheet, 46, 46, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 46, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 46, col);
}
//set empty cell for the 46th row(2nd set)
spreadsheetMergeCells(spreadsheet, 46, 46, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 46, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='thick',rightborder='thick'}, 46, col);
}
//set rhodium value for the 47th row
spreadsheetMergeCells(spreadsheet, 47, 47, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Rhodium(If required)', 47, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 47, 1);
spreadsheetMergeCells(spreadsheet, 47, 47, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 47, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 47, col);
}
//set empty cell for the 47th row(2nd set)
spreadsheetMergeCells(spreadsheet, 47, 47, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 47, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='thick',rightborder='thick'}, 47, col);
}
//set misc,texturing value for the 48th row
spreadsheetMergeCells(spreadsheet, 48, 48, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Misc(Texturing,Etc)', 48, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 48, 1);
spreadsheetMergeCells(spreadsheet, 48, 48, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 48, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 48, col);
}
//set empty cell for the 48th row(2nd set)
spreadsheetMergeCells(spreadsheet, 48, 48, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 48, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='thick',rightborder='thick'}, 48, col);
}
//set Set Center value for the 49th row
spreadsheetMergeCells(spreadsheet, 49, 49, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Set Center', 49, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 49, 1);
spreadsheetMergeCells(spreadsheet, 49, 49, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 49, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 49, col);
}
//set empty cell for the 49th row(2nd set)
spreadsheetMergeCells(spreadsheet, 49, 49, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 49, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='thick',rightborder='thick'}, 49, col);
}
//set Set Center value for the 50th row
spreadsheetMergeCells(spreadsheet, 50, 50, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Set Melee', 50, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 50, 1);
spreadsheetMergeCells(spreadsheet, 50, 50, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 50, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 50, col);
}
//set empty cell for the 50th row(2nd set)
spreadsheetMergeCells(spreadsheet, 50, 50, 8, 14);
spreadsheetSetCellValue(spreadsheet, '', 50, 8);
for(col=8;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thick',leftborder='thick',rightborder='thick'}, 50, col);
}
//set igi gia value for the 51st row
spreadsheetMergeCells(spreadsheet, 51, 51, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'IGI / GIA', 51, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 51, 1);
spreadsheetMergeCells(spreadsheet, 51, 51, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 51, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thick'}, 51, col);
}
//SET total labour information in the 52nd row
spreadsheetMergeCells(spreadsheet, 52, 52, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Total Labour', 52, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right',color='black',bold=true}, 52, 1);
spreadsheetMergeCells(spreadsheet, 52, 52, 4, 6);
spreadsheetSetCellValue(spreadsheet, '$0.00', 52, 4);
formatDollarValue=structNew();
formatDollarValue.dataFormat='$0.00';
formatDollarValue.alignment='center';
formatDollarValue.color='black';
formatDollarValue.bold=true;
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet,formatDollarValue , 52, col);
}
//set empty values in the 53rd row(1st set)
spreadsheetMergeCells(spreadsheet, 53, 53, 1, 2);
spreadsheetSetCellValue(spreadsheet, '', 53, 1);
spreadsheetFormatCell(spreadsheet, {}, 53, 1);
//set empty values in the 53rd row(2nd set)
spreadsheetMergeCells(spreadsheet, 53, 53, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 53, 4);
spreadsheetFormatCell(spreadsheet, {}, 53, 4);



// Set the content type and output the spreadsheet
</cfscript>
<cfheader name="Content-Disposition" value="inline; filename=#theFile#">
<cfcontent type="application/vnd.ms-excel" variable="#SpreadsheetReadBinary(spreadsheet)#">
