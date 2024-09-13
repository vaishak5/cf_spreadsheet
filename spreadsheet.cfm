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
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", fgcolor="light_yellow", color="black", alignment="center"}, 3, 1);
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
spreadsheetMergeCells(spreadsheet, 5, 5, 10, 12)
spreadsheetSetCellValue(spreadsheet, 'QUOTE PROVIDED BY NAME:', 5, 10);
formatQuote=structNew();
formatQuote.bold=true;
formatQuote.fontsize=12;
formatQuote.color='black';
formatQuote.alignment='right';
spreadsheetFormatCell(spreadsheet,formatQuote,5,10);
spreadsheetSetColumnWidth(spreadsheet, 10, 20);
spreadsheetSetColumnWidth(spreadsheet, 12, 20);
spreadsheetMergeCells(spreadsheet, 5, 5, 13, 16);
spreadsheetSetCellValue(spreadsheet, '', 5, 13);
for(col=13;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 5, col);
}
//set vendor/company name in the 6th row:
spreadsheetSetCellValue(spreadsheet, 'VENDOR / COMPANY NAME: ',  6, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black" ,alignment='right'},6,1);
//set borderbottom line for the vendor row
spreadsheetMergeCells(spreadsheet, 6, 6, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  6, 6);
for (col = 2; col <= 6; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 6, col);
    }
spreadsheetSetCellValue(spreadsheet, 'POSITION:',  6, 12);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black", alignment="right"},6,12);
spreadsheetMergeCells(spreadsheet, 6, 6, 13, 16);
spreadsheetSetCellValue(spreadsheet, '', 6, 13);
for(col=13;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 6, col);
}
//set ADDRESS name in the 7th row:
spreadsheetSetCellValue(spreadsheet, 'ADDRESS:',  7, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black",alignment="right"},7,1);
//set borderbottom line for the email row
spreadsheetMergeCells(spreadsheet, 7, 7, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  7, 6);
for (col = 2; col <= 6; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 7, col);
    }
spreadsheetSetCellValue(spreadsheet, 'EMAIL:',  7, 12);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black", alignment="right"},7,12);
spreadsheetMergeCells(spreadsheet, 7, 7, 13, 16);
spreadsheetSetCellValue(spreadsheet, '', 7, 13);
for(col=13;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 7, col);
}
//set CITY/STATE/ZIP name in the 8th row:
spreadsheetSetCellValue(spreadsheet, 'CITY / STATE / ZIP:',  8, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black",alignment='right'},8,1);
//set borderbottom line for the 8th row
spreadsheetMergeCells(spreadsheet, 8, 8, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  8, 6);
for (col = 2; col <= 6; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 8, col);
    }
spreadsheetSetColumnWidth(spreadsheet, 8, 15);
spreadsheetMergeCells(spreadsheet, 8, 8, 8, 12);
spreadsheetSetCellValue(spreadsheet, 'QUOTE IS VALID FOR WHICH COUNTRIES:', 8, 8);
spreadsheetFormatCell(spreadsheet, {bold=true,fontsize="12", color="black",alignment='right'}, 8, 8);
spreadsheetMergeCells(spreadsheet, 8, 8, 13, 16);
spreadsheetSetCellValue(spreadsheet, '', 8, 13);
for(col=13;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 8, col);
}
//set telephone name in the 9th row:
spreadsheetSetCellValue(spreadsheet, 'TELEPHONE ##:',  9, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},9,1);
//set borderbottom line for the 9th row
spreadsheetMergeCells(spreadsheet, 9, 9, 2, 6);
spreadsheetSetCellValue(spreadsheet, '',  9, 6);
for (col = 2; col <= 6; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 9, col);
}
// Set "ITEM INFORMATION" in the 11th row and apply format
spreadsheetMergeCells(spreadsheet, 11, 11, 1, 6);
spreadsheetSetCellValue(spreadsheet, "ITEM INFORMATION", 11, 1);
    spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="light_yellow", color="black", alignment="center"}, 11, 1);
for (col = 1; col <= 6; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin'}, 11, col);
}
//set PURCHASE ORDER INFORMATION in the 11th row
spreadsheetMergeCells(spreadsheet, 11, 11, 12, 16);
spreadsheetSetCellValue(spreadsheet, "PURCHASE ORDER INFORMATION", 11, 12);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="light_yellow", color="black", alignment="center", bgcolor="light_yellow"}, 11, 12);
for (col = 12; col <= 16; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin'}, 11, col);
}
spreadsheetSetColumnWidth(spreadsheet, 13, 13);
spreadsheetSetColumnWidth(spreadsheet, 16, 20);
//set dimensions  in the 12th row
spreadsheetSetCellValue(spreadsheet, 'Dimensions :',  12, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="right"},12,1);
spreadsheetSetColumnWidth(spreadsheet, 8, 15);
spreadsheetSetCellValue(spreadsheet, 'Height:', 12, 8);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black",alignment="center"}, 12, 8);
spreadsheetMergeCells(spreadsheet, 12, 12, 12, 14);
spreadsheetSetCellValue(spreadsheet, 'PURCHASE ORDER NUMBER', 12, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin',bgcolor="none",fontsize="12",alignment="center"}, 12, col);
}
spreadsheetMergeCells(spreadsheet, 12, 12, 15, 16);
spreadsheetSetCellValue(spreadsheet, 'QUANTITY & SHIP DATE', 12, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',bgcolor="none",fontsize="12",alignment="center"}, 12, col);
}
//set l,w,h,cube in the 13th row
spreadsheetSetCellValue(spreadsheet, 'L :',  13, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="right"},13,1);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 2);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 2);
spreadsheetSetCellValue(spreadsheet, 'W:',  13, 3);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="center"},13,3);
spreadsheetSetColumnWidth(spreadsheet, 3, 4);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 4);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 4);
spreadsheetSetCellValue(spreadsheet, 'H :',  13, 5);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="center"},13,5);
spreadsheetSetColumnWidth(spreadsheet, 5, 4);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 6);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 6);
spreadsheetSetCellValue(spreadsheet, 'H :',  13, 5);
spreadsheetFormatCell(spreadsheet, {fontsize="12", color="black", alignment="right"},13,5);
spreadsheetSetCellValue(spreadsheet, ' ', 13, 6);
spreadsheetFormatCell(spreadsheet, {bottomborder="thin"}, 13, 6);
spreadsheetSetCellValue(spreadsheet, 'Cube:', 13, 8);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black",alignment="center"}, 13, 8);
spreadsheetSetCellValue(spreadsheet, 0.00, 13, 10);
formatStruct=structNew();
formatStruct.alignment = "center";
formatStruct.fontsize = "12";
formatStruct.dataformat="0.00";
spreadsheetFormatCellRange(spreadsheet,formatStruct, 13, 10, 13, 10);
spreadsheetMergeCells(spreadsheet, 13, 13, 12, 14)
spreadsheetSetCellValue(spreadsheet, '', 13, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none"}, 13, col);
} 
spreadsheetMergeCells(spreadsheet, 13, 13, 15, 16)
spreadsheetSetCellValue(spreadsheet, '', 13, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none"}, 13, col);
} 
//set borderlines and purchase order number and ship date in the 14th row
spreadsheetMergeCells(spreadsheet, 14, 14, 12, 14);
spreadsheetSetCellValue(spreadsheet, 'PURCHASE ORDER NUMBER', 14, 12);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 14, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', rightborder='thin',leftborder='thin',bgcolor="none"}, 14, col);
}
spreadsheetMergeCells(spreadsheet,14, 14, 15, 16);
spreadsheetSetCellValue(spreadsheet, 'QUANTITY & SHIP DATE', 14, 15);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 14, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thin',bgcolor="none"}, 14, col);
} 
//set columnwidth for 7th,9th,11th cols
for (column = 7; column<12; column=column+2) {
        spreadSheetSetColumnWidth(spreadsheet, column, 2);        
    }
//set costco depost information in the 15th row
spreadsheetSetCellValue(spreadsheet, 'Costco Depot(889 / 894 / BOTH):', 15, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true}, 15, 1);
//set borderbottom line for the costco depost information(15th) row
spreadsheetMergeCells(spreadsheet, 15, 15, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  15, 2);
for (col = 2; col <= 8; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 15, col);
    } 
//SET EX: values in the 15th row
spreadsheetMergeCells(spreadsheet, 15, 15, 12, 14);
spreadsheetSetCellValue(spreadsheet, 'EX:8950101123', 15, 12);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 15, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', rightborder='thin',leftborder='thin',bgcolor="none"}, 15, col);
}
spreadsheetMergeCells(spreadsheet, 15, 15, 15, 16);
spreadsheetSetCellValue(spreadsheet, '11/1/31', 15, 15);
spreadsheetFormatCell(spreadsheet, {fontsize="12",alignment="center"}, 15, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='thin',bgcolor="none",alignment="center"}, 15, col);
}
//set reorder in the 16th row
spreadsheetSetCellValue(spreadsheet, 'Reorder(Y/N):', 16, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 16, 1);
spreadsheetMergeCells(spreadsheet, 16, 16, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  16, 2);
for (col = 2; col <= 8; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 16, col);
}

//set border line in the 16th row empty contents
spreadsheetMergeCells(spreadsheet, 16, 16, 12, 14)
spreadsheetSetCellValue(spreadsheet, '', 16, 12);
spreadsheetFormatCell(spreadsheet,{color="black"},16, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin"}, 16, col);
}
spreadsheetMergeCells(spreadsheet, 16, 16, 15, 16);
spreadsheetSetCellValue(spreadsheet, '', 16, 15);
spreadsheetFormatCell(spreadsheet, {color="black"},16, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 16, col);
} 
//set new item in the 17th row
spreadsheetSetCellValue(spreadsheet, 'New Item (Y/N):', 17, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 17, 1);
spreadsheetMergeCells(spreadsheet, 17, 17, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  17, 2);
for (col = 2; col <= 8; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 17, col);
    }
//set border line in the 17th row empty contents
spreadsheetMergeCells(spreadsheet, 17, 17, 12, 14)
spreadsheetSetCellValue(spreadsheet, '', 17, 12);
spreadsheetFormatCell(spreadsheet,{color="black"},17, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin"}, 17, col);
}
spreadsheetMergeCells(spreadsheet, 17, 17, 15, 16);
spreadsheetSetCellValue(spreadsheet, '', 17, 15);
spreadsheetFormatCell(spreadsheet, {color="black"},17, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 17, col);
} 
//set item description in the 18th row
spreadsheetSetCellValue(spreadsheet, 'Item Description:', 18, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 18, 1);
spreadsheetMergeCells(spreadsheet, 18, 18, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  18, 2);
for (col = 2; col <= 8; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 18, col);
    }
//set border line in the 18th row empty contents
spreadsheetMergeCells(spreadsheet, 18, 18, 12, 14)
spreadsheetSetCellValue(spreadsheet, '', 18, 12);
spreadsheetFormatCell(spreadsheet,{color="black"},18, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin"}, 18, col);
}
spreadsheetMergeCells(spreadsheet, 18, 18, 15, 16);
spreadsheetSetCellValue(spreadsheet, '', 18, 15);
spreadsheetFormatCell(spreadsheet, {color="black"},18, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 18, col);
} 
//set vendor style # in the 19th row
spreadsheetSetCellValue(spreadsheet, 'Vendor Style ##:', 19, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 19, 1);
spreadsheetMergeCells(spreadsheet, 19, 19, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  19, 2);
for (col = 2; col <= 8; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 19, col);
    }
//set border line in the 19th row empty contents
spreadsheetSetCellValue(spreadsheet, '', 19, 11);
spreadsheetFormatCell(spreadsheet, {rightborder='thin'}, 19, 11);
spreadsheetMergeCells(spreadsheet, 19, 19, 12, 14)
spreadsheetSetCellValue(spreadsheet, '', 19, 12);
formatBorders=structNew();
formatBorders.bgcolor="none";
formatBorders.bottomborder='thin';
formatBorders.rightborder="thin";
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, formatBorders, 19, col);
}

spreadsheetMergeCells(spreadsheet, 19, 19, 15, 16);
spreadsheetSetCellValue(spreadsheet, '', 19, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 19, col);
}  
//set minimum cwt in the 20th row
spreadsheetSetCellValue(spreadsheet, 'Minimum CWT:', 20, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 20, 1);
spreadsheetMergeCells(spreadsheet, 20, 20, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  20, 2);
for (col = 2; col <= 8; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 20, col);
}
//set border line in the 20th row empty contents
spreadsheetMergeCells(spreadsheet, 20,20, 12, 14)
spreadsheetSetCellValue(spreadsheet, '', 20, 12);
spreadsheetFormatCell(spreadsheet,{color="black"},19, 12);
for(col=12;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bgcolor="none",bottomborder='thin',rightborder="thin",topborder='thin'}, 20, col);
}
spreadsheetMergeCells(spreadsheet, 20, 20, 15, 16);
spreadsheetSetCellValue(spreadsheet, '', 20, 15);
spreadsheetFormatCell(spreadsheet, {color="black"},20, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='thin',bgcolor="none",bottomborder='thin'}, 20, col);
}   
//set minimum cwt in the 21st row
spreadsheetSetCellValue(spreadsheet, 'Minimum Center CWT:', 21, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 21, 1);
spreadsheetMergeCells(spreadsheet, 21, 21, 2, 8);
spreadsheetSetCellValue(spreadsheet, '',  21, 2);
for (col = 2; col <= 8; col++) {
        spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 21, col);
    }
//set costco item numbers in the 21st row
spreadsheetMergeCells(spreadsheet, 21, 21, 12, 16);
spreadsheetSetCellValue(spreadsheet, "COSTCO ITEM NUMBER(S)", 21, 12);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="light_yellow", color="black", alignment="center", bgcolor="light_yellow"}, 21, 12);
for (col = 12; col <= 16; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin'}, 21, col);
    
}
//set item/features information in the 22nd row
spreadsheetSetCellValue(spreadsheet, 'Item Features/Specs:', 22, 1);
spreadsheetFormatCell(spreadsheet, {fontsize="12",color='black',bold=true,alignment="right"}, 22, 1);
spreadsheetMergeCells(spreadsheet, 22, 22, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 22, col);
}
spreadsheetMergeCells(spreadsheet, 22, 22, 6, 8);
for(col=6;col<=8;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 22, col);
}
spreadsheetSetColumnWidth(spreadsheet, 15, 3);
//set 22nd row(secondset)
spreadsheetSetCellValue(spreadsheet, '##1', 22, 12);
spreadsheetFormatCell(spreadsheet, {leftborder='thin',alignment="right",bold=true}, 22, 12);
spreadsheetSetCellValue(spreadsheet, '##2', 22, 15);
spreadsheetFormatCell(spreadsheet, {color="black",bold=true},22, 15);
spreadsheetSetCellValue(spreadsheet, '', 22, 16);
spreadsheetFormatCell(spreadsheet, {rightborder='thin'}, 22, 16);
//set empty contents in the 23rd row
spreadsheetMergeCells(spreadsheet, 23, 23, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 23, col);
}
spreadsheetMergeCells(spreadsheet, 23, 23, 6, 8);
for(col=6;col<=8;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 23, col);
}
spreadsheetSetColumnWidth(spreadsheet, 15, 3);
//set 23 row(secondset)
spreadsheetSetCellValue(spreadsheet, '##3', 23, 12);
spreadsheetFormatCell(spreadsheet, {leftborder='thin',alignment="right",bold=true}, 23, 12);
spreadsheetSetCellValue(spreadsheet, '##4', 23, 15);
spreadsheetFormatCell(spreadsheet, {color="black",bold=true},23, 15);
spreadsheetSetCellValue(spreadsheet, '', 23, 16);
spreadsheetFormatCell(spreadsheet, {rightborder='thin'}, 23, 16); 
//set empty contents in the 24th row
spreadsheetMergeCells(spreadsheet, 24, 24, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 24, col);
}
spreadsheetMergeCells(spreadsheet, 24, 24, 6, 8);
for(col=6;col<=8;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 24, col);
}
spreadsheetSetColumnWidth(spreadsheet, 15, 3);
//set 23 row(secondset)
spreadsheetSetCellValue(spreadsheet, '##5', 24, 12);
spreadsheetFormatCell(spreadsheet, {leftborder='thin',alignment="right",bold=true,bottomborder='thin'}, 24, 12);
spreadsheetMergeCells(spreadsheet, 24, 24, 13, 14);
spreadsheetSetCellValue(spreadsheet, '', 24, 13);
for(col=13;col<=14;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 24, col);
}
spreadsheetSetCellValue(spreadsheet, '##6', 24, 15);
spreadsheetFormatCell(spreadsheet, {color="black",bold=true,bottomborder='thin'},24, 15);
spreadsheetSetCellValue(spreadsheet, '', 24, 16);
spreadsheetFormatCell(spreadsheet, {rightborder='thin',bottomborder='thin'}, 24, 16); 
//set empty contents in the 25th row
spreadsheetMergeCells(spreadsheet, 25, 25, 2, 4);
for(col=2;col<=4;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 25, col);
}
spreadsheetMergeCells(spreadsheet, 25, 25, 6, 8);
for(col=6;col<=8;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 25, col);
}
//set image information in the 26th row(secondset)
spreadsheetSetCellValue(spreadsheet, 'IMAGE:', 26, 12);
spreadsheetFormatCell(spreadsheet, {alignment='left'}, 26, 12);
//set item cost details in the 27th row 
spreadsheetMergeCells(spreadsheet, 27, 27, 1, 6);
spreadsheetSetCellValue(spreadsheet, "ITEM COST DETAILS", 27, 1);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="light_yellow", color="black", alignment="center"}, 27, 1);
for (col = 1; col <= 6; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='medium',rightborder='medium'}, 27, col);
}
//SET empty image set in the 27th-35th row
spreadsheetMergeCells(spreadsheet, 27, 35, 12, 16);
spreadsheetSetCellValue(spreadsheet, '', 27, 12);
spreadsheetFormatCell(spreadsheet, {bgcolor='none'}, 27, 12);
for (col = 12; col <= 16; col++) {
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
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='medium'}, 28, col);
}
//set usmca data details in the 29th row
spreadsheetSetCellValue(spreadsheet, "USMCA APPLICABLE (Y/N):", 29, 1);
spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", color="black", alignment="right"}, 29, 1);
spreadsheetSetCellValue(spreadsheet, '', 29, 2);
spreadsheetMergeCells(spreadsheet, 29, 29, 2, 6);
for(col=2;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='medium'}, 29, col);
}
//SET EMPTY CONTENT IN THE 30th row
spreadsheetSetCellValue(spreadsheet, "", 30, 1);
spreadsheetFormatCell(spreadsheet, {}, 30, 1);
spreadsheetSetCellValue(spreadsheet, '', 30, 2);
spreadsheetMergeCells(spreadsheet, 30, 30, 2, 6);
for(col=2;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='medium'}, 30, col);
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
        spreadsheetFormatCell(spreadsheet, {leftborder='thin',bottomborder='thin',rightborder='medium'}, 31, col);
    }
//set platinum content in the 32nd row
spreadsheetMergeCells(spreadsheet, 32, 32, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Platinum:', 32, 1);
spreadsheetFormatCell(spreadsheet, {alignment="right"}, 32, 1);
spreadsheetMergeCells(spreadsheet, 32, 32, 3, 6)
spreadsheetSetCellValue(spreadsheet, '', 32, 3);
for(col=3;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='medium',bottomborder='thin',leftborder='thin'}, 32, col);
}
//set minimum cwt in the 33rd row
spreadsheetMergeCells(spreadsheet, 33, 33, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Minimum CWT:', 33, 1);
spreadsheetFormatCell(spreadsheet, {alignment="right"}, 33, 1);
spreadsheetMergeCells(spreadsheet, 33, 33, 3, 6)
spreadsheetSetCellValue(spreadsheet, '', 33, 3);
for(col=3;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='medium',bottomborder='thin',leftborder='thin'}, 33, col);
}
//SET EMPTY contents in the 34th row
spreadsheetMergeCells(spreadsheet, 34, 34, 1, 2);
spreadsheetSetCellValue(spreadsheet, '', 34, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder="medium"}, 34, col);
}

spreadsheetMergeCells(spreadsheet, 34, 34, 3, 6)
spreadsheetSetCellValue(spreadsheet, '', 34, 3);
for(col=3;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {rightborder='medium',bottomborder='medium',leftborder='thin'}, 34, col);
}
//set mounting information in the 35th row
spreadsheetMergeCells(spreadsheet, 35, 35, 1, 6);
spreadsheetSetCellValue(spreadsheet, 'MOUNTING:', 35, 1);
spreadsheetformatcell(spreadsheet,{alignment="left",bold=true,color="black"},35,1);
for(col=1;col<=6;col++){
    spreadsheetformatcell(spreadsheet,{rightborder='medium'},35,col);
}
//set finished dwt in 36th row
spreadsheetMergeCells(spreadsheet, 36, 36, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Finished DWT', 36, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 36, 1);
spreadsheetMergeCells(spreadsheet, 36, 36, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 36, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 36, col);
}
//set casting charge in the 37th row
spreadsheetMergeCells(spreadsheet, 37, 37, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Casting Charge', 37, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 37, 1);
spreadsheetMergeCells(spreadsheet, 37, 37, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 37, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 37, col);
}
//set gold breakdown title in the 37th row(secondSet)
spreadsheetMergeCells(spreadsheet, 37, 37, 10, 12);
spreadsheetSetCellValue(spreadsheet, 'Gold Breakdown', 37, 10);
formatGoldValue=structNew();
formatGoldValue.color='black';
formatGoldValue.bold=true;
formatGoldValue.alignment='center';
formatGoldValue.fontsize='14';
formatGoldValue.bgcolor='yellow';
formatGoldValue.fgcolor='yellow';
formatGoldValue.topborder='medium';
formatGoldValue.bottomborder='thin';
formatGoldValue.rightborder='medium';
formatGoldValue.leftborder='medium';
for(col=10;col<=12;col++){
    spreadsheetFormatCell(spreadsheet, formatGoldValue, 37, col);
}
//set finding/chain in the 38th row
spreadsheetMergeCells(spreadsheet, 38, 38, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Finding / Chain', 38, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 38, 1);
spreadsheetMergeCells(spreadsheet, 38, 38, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 38, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 38, col);
}
//set gram informaion in the 38th row(secondset)
spreadsheetSetCellValue(spreadsheet, 'Gram:', 38, 10);
spreadsheetFormatCell(spreadsheet, {alignment='left',bgcolor='yellow',fgcolor='yellow',fontsize='13',leftborder='medium',rightborder='thin',bottomborder='thin'}, 38, 10);
spreadsheetMergeCells(spreadsheet, 38, 38, 11, 12);
spreadsheetSetCellValue(spreadsheet, '', 38, 11);
for(col=11;col<=12;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 38, col);
}
//set pcs on Casting in the 39th row
spreadsheetMergeCells(spreadsheet, 39, 39, 1, 2);
spreadsheetSetCellValue(spreadsheet, '## Pcs on Casting', 39, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 39, 1);
spreadsheetMergeCells(spreadsheet, 39, 39, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 39, 4);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 39, 1);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 39, col);
}
//set labour informaion in the 39th row(secondset)
spreadsheetSetCellValue(spreadsheet, 'Labour:', 39, 10);
spreadsheetFormatCell(spreadsheet, {alignment='left',bgcolor='yellow',fgcolor='yellow',fontsize='13',leftborder='medium',rightborder='thin',bottomborder='thin'}, 39, 10);
spreadsheetMergeCells(spreadsheet, 39, 39, 11, 12);
spreadsheetSetCellValue(spreadsheet, '', 39, 11);
for(col=11;col<=12;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 39, col);
}
//set head size/shape on Casting in the 40th row
spreadsheetMergeCells(spreadsheet, 40, 40, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Head Size / Shape', 40, 1);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 40, 1);
spreadsheetMergeCells(spreadsheet, 40, 40, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 40, 4);
spreadsheetFormatCell(spreadsheet, {alignment='right'}, 40, 1);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 40, col);
}
//set $ per gram informaion in the 40th row(secondset)
spreadsheetSetCellValue(spreadsheet, '$ Per Gram', 40, 10);
spreadsheetFormatCell(spreadsheet, {alignment='left',bgcolor='yellow',fgcolor='yellow',fontsize='13',leftborder='medium',rightborder='thin',bottomborder='medium'}, 40, 10);
spreadsheetMergeCells(spreadsheet, 40, 40, 11, 12);
spreadsheetSetCellValue(spreadsheet, '$0.00', 40, 11);
for(col=11;col<=12;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium',alignment='center',bottomborder='medium'}, 40, col);
}
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
formatValue.rightborder="medium";
for(col=4;col<=6;col++){
    spreadsheetFormatCellRange(spreadsheet,formatValue, 41, 4, 41, 6);
}
//set empty bottom border line in the 42nd row
spreadsheetMergeCells(spreadsheet, 42, 42, 1, 6)
spreadsheetSetCellValue(spreadsheet, '', 42, 1);
for(col=1;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium',rightborder='medium'}, 42, col);
}
//set labor cost details in the 43rd row(1st set)
spreadsheetMergeCells(spreadsheet, 43, 43, 1, 6);
spreadsheetSetCellValue(spreadsheet, 'LABOR COSTS :', 43, 1);
for(col=1;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='left',fontsize='12',bold='true',color='black',rightborder='medium'}, 43, col)
}
//set labor cost details in the 43rd row(2nd set)
spreadsheetMergeCells(spreadsheet, 43, 43, 10, 16);
spreadsheetSetCellValue(spreadsheet, 'LABOR COST DETAILS ', 43, 10);
formatLaborValue=structNew();
formatLaborValue.alignment='center';
formatLaborValue.bgcolor='gold';
formatLaborValue.fgcolor='gold';
formatLaborValue.topborder='medium';
formatLaborValue.bottomborder='medium';
formatLaborValue.rightborder='medium';
formatLaborValue.leftborder='medium';
formatLaborValue.fontsize='13';
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, formatLaborValue, 43, col);
}
//set cost to assemble value for the 44th row(cost to assemble)
spreadsheetMergeCells(spreadsheet, 44, 44, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Cost To Assemble', 44, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 44, 1);
spreadsheetMergeCells(spreadsheet, 44, 44, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 44, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 44, col);
}
//set empty cell for the 44th row(second set)
spreadsheetMergeCells(spreadsheet, 44, 44, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 44, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='medium',rightborder='medium'}, 44, col);
}
//set what needs to be assembled value for the 45th row
spreadsheetMergeCells(spreadsheet, 45, 45, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'What Needs to be Assembled', 45, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 45, 1);
spreadsheetMergeCells(spreadsheet, 45, 45, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 45, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 45, col);
}
//set empty cell for the 45th row(2nd set)
spreadsheetMergeCells(spreadsheet, 45, 45, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 45, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='medium',rightborder='medium'}, 45, col);
}
//set polish and finish value for the 46th row
spreadsheetMergeCells(spreadsheet, 46, 46, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Polish & Finish', 46, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 46, 1);
spreadsheetMergeCells(spreadsheet, 46, 46, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 46, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 46, col);
}
//set empty cell for the 46th row(2nd set)
spreadsheetMergeCells(spreadsheet, 46, 46, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 46, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='medium',rightborder='medium'}, 46, col);
}
//set rhodium value for the 47th row
spreadsheetMergeCells(spreadsheet, 47, 47, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Rhodium(If required)', 47, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 47, 1);
spreadsheetMergeCells(spreadsheet, 47, 47, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 47, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 47, col);
}
//set empty cell for the 47th row(2nd set)
spreadsheetMergeCells(spreadsheet, 47, 47, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 47, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='medium',rightborder='medium'}, 47, col);
}
//set misc,texturing value for the 48th row
spreadsheetMergeCells(spreadsheet, 48, 48, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Misc(Texturing,Etc)', 48, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 48, 1);
spreadsheetMergeCells(spreadsheet, 48, 48, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 48, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 48, col);
}
//set empty cell for the 48th row(2nd set)
spreadsheetMergeCells(spreadsheet, 48, 48, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 48, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='medium',rightborder='medium'}, 48, col);
}
//set Set Center value for the 49th row
spreadsheetMergeCells(spreadsheet, 49, 49, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Set Center', 49, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 49, 1);
spreadsheetMergeCells(spreadsheet, 49, 49, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 49, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 49, col);
}
//set empty cell for the 49th row(2nd set)
spreadsheetMergeCells(spreadsheet, 49, 49, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 49, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',leftborder='medium',rightborder='medium'}, 49, col);
}
//set Set Center value for the 50th row
spreadsheetMergeCells(spreadsheet, 50, 50, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Set Melee', 50, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 50, 1);
spreadsheetMergeCells(spreadsheet, 50, 50, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 50, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 50, col);
}
//set empty cell for the 50th row(2nd set)
spreadsheetMergeCells(spreadsheet, 50, 50, 10, 16);
spreadsheetSetCellValue(spreadsheet, '', 50, 10);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium',leftborder='medium',rightborder='medium'}, 50, col);
}
//set igi gia value for the 51st row
spreadsheetMergeCells(spreadsheet, 51, 51, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'IGI / GIA', 51, 1);
spreadsheetFormatCell(spreadsheet, {alignment: 'right'}, 51, 1);
spreadsheetMergeCells(spreadsheet, 51, 51, 4, 6);
spreadsheetSetCellValue(spreadsheet, '', 51, 4);
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin',rightborder='medium'}, 51, col);
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
formatDollarValue.rightborder='medium';
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
for(col=4;col<=6;col++) {
    spreadsheetFormatCell(spreadsheet, {rightborder='medium'}, 53, col);
}
//set total section value in the 54th row(1st set)
spreadsheetMergeCells(spreadsheet, 54, 54, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'TOTAL SECTION 1', 54, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right',color='black',bold=true,bottomborder='medium'}, 54, col);
}
spreadsheetSetCellValue(spreadsheet, '', 54, 3);
spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 54, 3);

spreadsheetMergeCells(spreadsheet, 54, 54, 4, 6);
spreadsheetSetCellValue(spreadsheet, '$0.00', 54, 4);
formatDollarValue=structNew();
formatDollarValue.dataFormat='$0.00';
formatDollarValue.alignment='center';
formatDollarValue.color='black';
formatDollarValue.bold=true;
formatDollarValue.rightborder='medium';
formatDollarValue.fgcolor='grey_40_percent';
formatDollarValue.bottomborder='medium';
formatDollarValue.topborder='thin';
for(col=4;col<=6;col++){
    spreadsheetFormatCell(spreadsheet,formatDollarValue , 54, col);
}
//set diamond breakdown value in the 56th row(1st set)
spreadsheetSetCellValue(spreadsheet, 'DIAMOND BREAKDOWN :', 56, 1);
formatDiamondSet=structNew();
formatDiamondSet.color='black';
formatDiamondSet.fgcolor='light_yellow';
formatDiamondSet.bold=true;
formatDiamondSet.alignment='right';
formatDiamondSet.topborder='medium';
formatDiamondSet.bottomborder='medium';
spreadsheetFormatCell(spreadsheet, formatDiamondSet, 56, 1);
//set empty in the 56th row(2nd set)
spreadsheetMergeCells(spreadsheet, 56, 56, 2, 16);
spreadsheetSetCellValue(spreadsheet, '', 56, 2);
formatEmptyVal=structNew();
formatEmptyVal.rightborder='medium';
formatEmptyVal.topborder='medium';
formatEmptyVal.fgcolor='light_yellow';
formatEmptyVal.bottomborder='medium';
for(col=2;col<=16;col++){
    spreadsheetFormatCell(spreadsheet,formatEmptyVal , 56, col);
}
spreadsheetSetCellValue(spreadsheet, 'QTY', 57, 2);
formatQTY=structNew();
formatQTY.alignment='center';
formatQTY.bold=true;
formatQTY.fontsize=12;
formatQTYLast=structNew();
formatQTYLast.alignment='center';
formatQTYLast.bold=true;
formatQTYLast.fontsize=12;
formatQTYLast.rightborder='medium';
spreadsheetFormatCell(spreadsheet, formatQTY, 57, 2);//2nd column
spreadsheetSetCellValue(spreadsheet, 'WT EA', 57, 4);//4th column
spreadsheetFormatCell(spreadsheet, formatQTY, 57, 4);
spreadsheetSetColumnWidth(spreadsheet, 6, 15);
spreadsheetSetCellValue(spreadsheet, 'BILLED TWT', 57, 6);//6th column
spreadsheetFormatCell(spreadsheet, formatQTY, 57, 6);
spreadsheetSetCellValue(spreadsheet, '$ PER CT', 57, 8);//8th column
spreadsheetFormatCell(spreadsheet, formatQTY, 57, 8);
spreadsheetMergeCells(spreadsheet, 57, 57, 10, 13);
spreadsheetSetCellValue(spreadsheet, 'SHAPE & MINIMUM WT ETA', 57, 10);//10th-13th column
for(col=10;col<=13;col++){
    spreadsheetFormatCell(spreadsheet, formatQTY, 57, col);
}
spreadsheetMergeCells(spreadsheet, 57, 57, 15, 16);
spreadsheetSetCellValue(spreadsheet, 'EXIT PRICE', 57, 15);//15-16 column
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, formatQTYLast, 57, col);
}
//SET values in the 58th row
spreadsheetSetCellValue(spreadsheet, 'Center Diamond', 58, 1);
formatCenterDiamond=structNew();
formatCenterDiamond.alignment='right';
formatCenterDiamond.fontsize=12;
spreadsheetFormatCell(spreadsheet, formatCenterDiamond, 58, 1);
spreadsheetSetCellValue(spreadsheet, '', 58, 2);
formatEmptyCells=structNew();
formatEmptyCells.bottomborder='thin';
formatEmptyCells.alignment='center';
formatEmptyCellsLast=structNew();
formatEmptyCellsLast.bottomborder='thin';
formatEmptyCellsLast.alignment='center';
formatEmptyCellsLast.rightborder='medium';
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 58, 2);
spreadsheetSetCellValue(spreadsheet, '', 58, 4);
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 58, 4);
spreadsheetSetCellValue(spreadsheet, '0', 58, 6);
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 58, 6);
spreadsheetSetCellValue(spreadsheet, '##DIV/01!', 58, 8);
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 58, 8);
spreadsheetMergeCells(spreadsheet, 58, 58, 10, 13);
spreadsheetSetCellValue(spreadsheet, '', 58, 10);
for(col=10;col<=13;col++){
    spreadsheetFormatCell(spreadsheet, formatEmptyCells, 58, col);
}
spreadsheetMergeCells(spreadsheet, 58, 58, 15, 16);
spreadsheetSetCellValue(spreadsheet, '', 58, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, formatEmptyCellsLast, 58, col);
}
//set Melee value in the 59th to 65th row
formatCenterDiamond = structNew();
formatCenterDiamond.alignment = 'right';
formatCenterDiamond.fontsize = 12;
formatEmptyCells = structNew();
formatEmptyCells.bottomborder = 'thin';
formatEmptyCells.alignment = 'center';
formatLastCell=structNew();
formatLastCell.rightborder = 'medium';
formatLastCell.bottomborder = 'thin';
//set loop for the start rwo to end row
for(row=59; row<=65; row++) {
    spreadsheetSetCellValue(spreadsheet, 'Melee', row, 1);
    spreadsheetFormatCell(spreadsheet, formatCenterDiamond, row, 1);
    spreadsheetSetCellValue(spreadsheet, '', row, 2);
    spreadsheetFormatCell(spreadsheet, formatEmptyCells, row, 2);
    spreadsheetSetCellValue(spreadsheet, '', row, 4);
    spreadsheetFormatCell(spreadsheet, formatEmptyCells, row, 4);
    spreadsheetSetCellValue(spreadsheet, '0', row, 6);
    spreadsheetFormatCell(spreadsheet, formatEmptyCells, row, 6);
    spreadsheetSetCellValue(spreadsheet, '##DIV/01!', row, 8);
    spreadsheetFormatCell(spreadsheet, formatEmptyCells, row, 8);
    spreadsheetMergeCells(spreadsheet, row, row, 10, 13);
    spreadsheetSetCellValue(spreadsheet, '', row, 10);
    for(col=10; col<=13; col++) {
        spreadsheetFormatCell(spreadsheet, formatEmptyCells, row, col);
    }
    spreadsheetMergeCells(spreadsheet, row, row, 15, 16);
    spreadsheetSetCellValue(spreadsheet, '', row, 15);
    for(col=15; col<=16; col++) {
        spreadsheetFormatCell(spreadsheet, formatLastCell, row, col);
    }
}
//set value qty in the 66th row
spreadsheetSetCellValue(spreadsheet, 'QTY', 66, 1);
formatCenterDiamond=structNew();
formatCenterDiamond.alignment='right';
formatCenterDiamond.fontsize=12;
formatCenterDiamond.bottomborder='medium';
spreadsheetFormatCell(spreadsheet, formatCenterDiamond, 66, 1);
spreadsheetSetCellValue(spreadsheet, '0', 66, 2);
formatEmptyCells=structNew();
formatEmptyCells.bottomborder='medium';
formatEmptyCells.alignment='center';
formatSecondCell=structNew();
formatSecondCell.bottomborder='medium';
formatSecondCell.alignment='center';
formatSecondCell.dataformat='0.00'
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 66, 2);
spreadsheetSetCellValue(spreadsheet, '', 66, 4);
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 66, 4);
spreadsheetSetCellValue(spreadsheet, '0.00', 66, 6);
spreadsheetFormatCell(spreadsheet, formatSecondCell, 66, 6);
spreadsheetFormatCell(spreadsheet, formatEmptyCells, 66, 8);
spreadsheetMergeCells(spreadsheet, 66, 66, 10, 13);
spreadsheetSetCellValue(spreadsheet, 'DIAMOND TOTAL SECTION 2: ', 66, 10);
for(col=10;col<=13;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right',bottomborder='medium',bold=true}, 66, col);
}
formatLast=structNew();
formatLast.bottomborder='medium';
formatLast.alignment='center';
formatLast.bold=true;
formatLast.fgcolor='grey_40_percent';
formatLast.rightborder='medium';
spreadsheetMergeCells(spreadsheet, 66, 66, 15, 16);
spreadsheetSetCellValue(spreadsheet, '$0.00', 66, 15);
for(col=15;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, formatLast, 66, col);
}
 for(col=3;col<=9;col=col+2){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 66, col);
} 
spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 66, 14);
//set diamond breakdown value in the 68th row(1st set)
spreadsheetSetCellValue(spreadsheet, 'COLOR BREAKDOWN:', 68, 1);
formatDiamondSet=structNew();
formatDiamondSet.color='black';
formatDiamondSet.fgcolor='light_yellow';
formatDiamondSet.bold=true;
formatDiamondSet.alignment='right';
formatDiamondSet.topborder='medium';
formatDiamondSet.bottomborder='medium';
spreadsheetFormatCell(spreadsheet, formatDiamondSet, 68, 1);
//set empty in the 68th row(2nd set)
spreadsheetMergeCells(spreadsheet, 68, 68, 2, 16);
spreadsheetSetCellValue(spreadsheet, '', 68, 2);
formatEmptyVal=structNew();
formatEmptyVal.rightborder='medium';
formatEmptyVal.topborder='medium';
formatEmptyVal.fgcolor='light_yellow';
formatEmptyVal.bottomborder='medium';
for(col=2;col<=16;col++){
    spreadsheetFormatCell(spreadsheet,formatEmptyVal , 68, col);
}
//set contents in the 69th row
spreadsheetSetCellValue(spreadsheet, 'QTY', 69, 2);
formatQTY=structNew();
formatQTY.alignment='center';
formatQTY.bold=true;
formatQTY.fontsize=12;
formatEmptyCellsLastCol = structNew();
formatEmptyCellsLastCol.bottomborder = 'thin';
formatEmptyCellsLastCol.alignment = 'center';
formatEmptyCellsLastCol.rightborder='medium';
formatEmptyCellsLastCol.bold=true;
spreadsheetFormatCell(spreadsheet, formatQTY, 69, 2);//2nd column
spreadsheetSetCellValue(spreadsheet, 'WT EA', 69, 4);//4th column
spreadsheetFormatCell(spreadsheet, formatQTY, 69, 4);
spreadsheetSetColumnWidth(spreadsheet, 6, 15);
spreadsheetSetCellValue(spreadsheet, 'SHAPE', 69, 6);//6th column
spreadsheetFormatCell(spreadsheet, formatQTY, 69, 6);
spreadsheetSetCellValue(spreadsheet, 'MM SIZES', 69, 8);//8th column
spreadsheetFormatCell(spreadsheet, formatQTY, 69, 8);
spreadsheetSetCellValue(spreadsheet, '$ PER CT', 69, 10);//10th column
spreadsheetFormatCell(spreadsheet, formatQTY, 69, 10);
spreadsheetSetCellValue(spreadsheet, 'EXT PRICE', 69, 12);//12th column
spreadsheetFormatCell(spreadsheet, formatQTY, 69, 12);
spreadsheetSetCellValue(spreadsheet, 'COLOR', 69, 16);//16th column
spreadsheetFormatCell(spreadsheet, formatEmptyCellsLastCol, 69, 16);
//set contents in the 70th row to 74th row
formatColor=structNew();
formatColor.alignment='right';
formatEmptyCells = structNew();
formatEmptyCells.bottomborder = 'thin';
formatEmptyCells.alignment = 'center';
for(row=70;row<=74;row++){
    spreadsheetSetCellValue(spreadsheet, 'Color',row, 1);
    spreadsheetFormatCell(spreadsheet, formatColor, row, 1);
    for(col=2;col<=10;col=col+2){
        spreadsheetFormatCell(spreadsheet, formatEmptyCells, row, col);
    }
    spreadsheetSetCellValue(spreadsheet, "$0.00", row, 12);
    spreadsheetFormatCell(spreadsheet, formatEmptyCells, row,12);
    spreadsheetFormatCell(spreadsheet, formatEmptyCellsLastCol, row, 16);
}
//set colour section contents in the 75th row
spreadsheetMergeCells(spreadsheet, 75, 75, 1, 7);
for(col=1;col<=7;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 75, col);
}
spreadsheetMergeCells(spreadsheet, 75, 75, 8, 10);
spreadsheetSetCellValue(spreadsheet, 'COLOUR TOTAL SECTION 3:', 75,8);
formatColour=structNew();
formatColour.alignment='right';
formatColour.bold=true;
formatColour.color='black';
formatColour.bottomborder='medium';
for(col=8;col<=10;col++){
    spreadsheetFormatCell(spreadsheet, formatColour, 75, col);
}
spreadsheetFormatCell(spreadsheet, {bottomborder='medium'}, 75, 11);
formatDollarCont=structNew();
formatDollarCont.bold=true;
formatDollarCont.alignment='center';
formatDollarCont.leftborder='thin';
formatDollarCont.fgcolor='pale_blue';
formatDollarCont.bottomborder='medium';
spreadsheetSetCellValue(spreadsheet, '$0.00', 75, 12);
spreadsheetFormatCell(spreadsheet, formatDollarCont, 75, 12);
spreadsheetMergeCells(spreadsheet, 75, 75, 13, 16);
formatLast=structNew();
formatLast.bottomborder='medium';
formatLast.rightborder='medium';
for(col=13;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, formatLast, 75, col);
}
//set values in the 77th row(1st set)
spreadsheetMergeCells(spreadsheet, 77, 77, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'TOTAL SECTIONS 1-3:', 77, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {underline='true',bold=true,alignment='right',fontsize=13}, 77, 1);
}
spreadsheetMergeCells(spreadsheet, 77, 77, 4, 5);
spreadsheetSetCellValue(spreadsheet, '$0.00', 77, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin'}, 77, col)
}
spreadsheetMergeCells(spreadsheet, 77, 77, 10, 16);
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 77, col);
}
//set contents in the 78th row
column=1;row=78;
spreadsheetSetCellValue(spreadsheet, 'Vendor Allowance %:', row, column+1);
spreadsheetFormatCellRange(spreadsheet, {alignment='right'}, row, column, row, column+1);
spreadsheetMergeCells(spreadsheet, 78, 78, 4, 5);
spreadsheetSetCellValue(spreadsheet, '2%', 78, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin'}, 78, col);
}
//set contents in the 79th row
column=1;row=79;
spreadsheetSetCellValue(spreadsheet, 'Vendor Allowance $:', row, column+1);
spreadsheetFormatCellRange(spreadsheet, {alignment='right'}, row, column, row, column+1);
spreadsheetMergeCells(spreadsheet, row, row, column+3, column+4);
spreadsheetSetCellValue(spreadsheet, '$0.00', row, column+3);
for(col=column+3;col<=column+4;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin'}, row, col);
}
//set contents in the 80th row
spreadsheetMergeCells(spreadsheet, 80, 80, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Marketing Allowance %:', 80, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 80, 1);
}
spreadsheetMergeCells(spreadsheet, 80, 80, 4, 5);
spreadsheetSetCellValue(spreadsheet, '0.50%', 80, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin'}, 80, col);
}
//set contents in the 81sth row
spreadsheetMergeCells(spreadsheet, 81, 81, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Marketing Allowance $:', 81, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 81, 1);
}
spreadsheetMergeCells(spreadsheet, 81, 81, 4, 5);
spreadsheetSetCellValue(spreadsheet, '$0.00', 81, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin'}, 81, col);
}
//set contents in the 82nd row
column=1;row=82;
spreadsheetSetCellValue(spreadsheet, 'Spoils Allowance %:', row, column+1);
spreadsheetFormatCellRange(spreadsheet, {alignment='right'}, row, column, row, column+1);
//set contents in the 83rd row
spreadsheetMergeCells(spreadsheet, 83, 83, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Sub Total:', 83, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 83, col);
}
spreadsheetMergeCells(spreadsheet, 83, 83, 4, 5);
spreadsheetSetCellValue(spreadsheet, '$0.00', 83, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin',topborder='thin',rightborder='thin',leftborder='thin',fgcolor='grey_40_percent'}, 83, col);
}
//set contents in the 84th row
spreadsheetMergeCells(spreadsheet, 84, 84, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Total DFI %(No Spoils):', 84, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 84, col);
}
spreadsheetMergeCells(spreadsheet, 84, 84, 4, 5);
spreadsheetSetCellValue(spreadsheet, 'N/A', 84, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin',topborder='thin',rightborder='thin',leftborder='thin',fgcolor='gold'}, 84, col);
}
//set contents in the 85th row
spreadsheetMergeCells(spreadsheet, 85, 85, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Total NET Cost:', 85, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 85, col);
}
spreadsheetMergeCells(spreadsheet, 85, 85, 4, 5);
spreadsheetSetCellValue(spreadsheet, '$0.00', 85, 4);
for(col=4;col<=5;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bottomborder='thin',topborder='thin',rightborder='thin',leftborder='thin',fgcolor='pale_blue'}, 85, col);
}
//set contents in the 87th row
column=1;row=87;
spreadsheetSetCellValue(spreadsheet, 'Socialized Casting:', row, column+1);
spreadsheetFormatCellRange(spreadsheet, {alignment='right',bold=true,underline='true'}, row, column, row, column+1);
spreadsheetMergeCells(spreadsheet, row, row, column+3, column+4);
spreadsheetFormatCellRange(spreadsheet, {bottomborder='thin'}, row, column+3, row, column+4);
//set contents in the 88-89row
startRow = 88;
startColumn = 1;
for(row = startRow; row <= startRow + 1; row++) {
    spreadsheetSetCellValue(spreadsheet, 'Finished DWT XXX:', row, startColumn + 1);
    spreadsheetFormatCellRange(spreadsheet, {alignment: 'right'}, row, startColumn, row, startColumn + 1);
}
//set contents in the 90,91strow
spreadsheetMergeCells(spreadsheet, 90, 90, 1, 2);
spreadsheetMergeCells(spreadsheet, 91, 91, 1, 2);
for(row=90;row<=91;row++){
    for(col=1;col<=2;col++){
            spreadsheetSetCellValue(spreadsheet, 'Finished DWT XXX:', row,col);
            spreadsheetFormatCell(spreadsheet, {alignment='right'}, row, col);
    }
}
//set border line in the 88-92 row
for(row=88;row<=92;row++){
    for(col=4;col<=5;col++){
        spreadsheetformatcell(spreadsheet,{bottomborder='thin'},row,col);
    }
}
//set content in the 92nd row
spreadsheetMergeCells(spreadsheet, 92, 92, 1, 2);
spreadsheetSetCellValue(spreadsheet, 'Socialized Cost All Sizes:', 92, 1);
for(col=1;col<=2;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 92, col);
}
//set values in the 81st row (2nd set)
spreadsheetMergeCells(spreadsheet, 81, 81, 10, 16);
spreadsheetSetCellValue(spreadsheet, 'SIZING', 81, 10);;
for(col=10;col<=16;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='center',bold=true, topborder='thin',bottomborder='thin',leftborder='thin',rightborder='thin', fgcolor='light_yellow'}, 81, col);
}
//set values in the 82nd row
spreadsheetSetCellValue(spreadsheet, 'If a Ring,Is It Sizeable?', 82, 12);
spreadsheetFormatCell(spreadsheet, {}, 82, 12);
//set value in the 83rd row
spreadsheetMergeCells(spreadsheet, 83, 83, 10, 12);
spreadsheetSetCellValue(spreadsheet, 'If Yes,How Much Can Ring Be Sized', 83,10);
for(col=10;col<=12;col++){
    spreadsheetFormatCell(spreadsheet, {alignment='right'}, 83, col);
} 
//set values in the 84th row
spreadsheetSetCellValue(spreadsheet, 'Finished Cost 5:', 84, 12);
spreadsheetFormatCell(spreadsheet, {}, 84, 12);
//set border lines in the 83-87th row
for(row=83;row<=87;row++){
    for(col=13;col<=15;col++){
        spreadsheetformatcell(spreadsheet,{bottomborder='thin'},row,col);
    }
}
//Set the content type and output the spreadsheet
</cfscript>

<cfheader name="Content-Disposition" value="inline; filename=#theFile#">
<cfcontent type="application/vnd.ms-excel" variable="#SpreadsheetReadBinary(spreadsheet)#">
