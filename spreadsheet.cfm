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
// Set the content type and output the spreadsheet
</cfscript>
<cfheader name="Content-Disposition" value="inline; filename=#theFile#">
<cfcontent type="application/vnd.ms-excel" variable="#SpreadsheetReadBinary(spreadsheet)#">
