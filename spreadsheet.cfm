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
theFile = "#DateFormat(now(), 'mm-dd-YYYY')#.xlsx";
// Create a new spreadsheet
spreadsheet = spreadsheetNew("Gemstones_ Pearl Quote Sheet", true);
spreadsheetCreateSheet(spreadsheet, 'Diamond Quote Sheet');
spreadSheetSetActiveSheet(spreadsheet, 'Diamond Quote Sheet');
imagePath = expandPath("images/Costco_Logo.png"); 
spreadsheetAddImage(spreadsheet, imagePath, "1, 1, 2, 2");
spreadsheetSetColumnWidth(spreadsheet, 1, 30);
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
spreadsheetSetCellValue(spreadsheet, '',  5, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 5, col);
}
//set vendor/company name in the 6th row:
spreadsheetSetCellValue(spreadsheet, 'VENDOR / COMPANY NAME: ',  6, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},6,1);
//set borderbottom line for the vendor row
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
spreadsheetSetCellValue(spreadsheet, '',  6, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 6, col);
}
//set ADDRESS name in the 7th row:
spreadsheetSetCellValue(spreadsheet, 'ADDRESS:',  7, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black",alignment="right"},7,1);
//set borderbottom line for the email row
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
spreadsheetSetCellValue(spreadsheet, '',  7, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 7, col);
}

//set CITY/STATE/ZIP name in the 8th row:
spreadsheetSetCellValue(spreadsheet, 'CITY / STATE / ZIP:',  8, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},8,1);
//set borderbottom line for the 8th row
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
spreadsheetSetCellValue(spreadsheet, '',  8, 16);
for (col = 10; col <= 14; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin'}, 8, col);
}
//set telephone name in the 9th row:
spreadsheetSetCellValue(spreadsheet, 'TELEPHONE ##:',  9, 1);
spreadsheetFormatCell(spreadsheet, {bold="true", fontsize="12", color="black"},9,1);
//set borderbottom line for the 9th row
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
spreadsheetMergeCells(spreadsheet, 11, 11, 9, 10);
spreadsheetSetCellValue(spreadsheet, "PURCHASE ORDER INFORMATION", 11, 9);
    spreadsheetFormatCell(spreadsheet, {bold=true, fontsize="12", fgcolor="yellow", color="black", alignment="center", bgcolor="yellow"}, 11, 9);
for (col = 9; col <= 10; col++) {
    spreadsheetFormatCell(spreadsheet, {bottomborder='thin', topborder='thin',rightborder='thin',leftborder='thin'}, 11, col);
    
}
// Set the content type and output the spreadsheet
</cfscript>
<cfheader name="Content-Disposition" value="inline; filename=#theFile#">
<cfcontent type="application/vnd.ms-excel" variable="#SpreadsheetReadBinary(spreadsheet)#">
