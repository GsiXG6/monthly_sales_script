const g_Color_Default	= "#BFBFBF";
const g_Color_White		= "#FFFFFF";
const g_Color_Black		= "#000000";

const g_iHeaderNumb		= 2;
const g_iBodyNumb		= 120;
const g_iFooterNumb		= 3;
const g_iSaleHeader		= g_iHeaderNumb + 1;
const g_iBodyStart		= g_iSaleHeader + 1;
const g_iBodyEnd		= g_iSaleHeader + g_iBodyNumb;
const g_iFootStart		= g_iBodyEnd + 1;
const g_iFootEnd		= g_iBodyEnd + g_iFooterNumb;

const CODE_COMPLETED	= '"F"';
const CODE_PENDING		= '"P"';
const CODE_HAVESTOCK	= '"HS"';
const CODE_FORWARD		= '"FN"';
const CODE_CANCLED		= '"C"';
const CODE_URGENT		= 3;
const CODE_NEARDUE		= 7;

const COLOR_COMPLETED	= "#00B0F0";
const COLOR_PENDING		= "#FF8080";
const COLOR_HAVESTOCK	= "green";
const COLOR_FORWARD		= "#7030A0";
const COLOR_CANCLED		= "grey";
const COLOR_URGENT		= "red";
const COLOR_NEARDUE		= "#E26B0A";

const g_bResetCellValue = false;
const g_bResetFillColor = true;
const g_bResetConFormat = true;
const g_bRunActiveOnly	= false;

let g_sLetterBuff0		=  "";
let g_sLetterBuff1		=  "";
let g_sLetterBuff2		=  "";
let g_sLetterBuff3		=  "";
let g_sLetterTemps1		=  "";
let g_sLetterTemps2		=  "";

const g_sSheetName =
[
	"Jan",
	"Feb",
	"Mar",
	"Apr",
	"May",
	"Jun",
	"Jul",
	"Aug",
	"Sep",
	"Oct",
	"Nov",
	"Dec",
];

const g_sDescName =
[
	"MKM DO/PO",
	"JOB No.",
	"CUSTOMER PO",
	"CODE",
	"DRAWING No.",
	"REV",
	"DESCRIPTION",
	"D/RECEIVED",
	"D/DUE",
	"QTY",
	"U/PRICE",
	"STATUS",
	"DAY DUE",
	"TOTAL SALES",
	"ACTUAL SALES",
	"D/DELIVER",
	"COMMENT"
];

const g_iDescSize =
[
	130,
	72,
	72,
	33,
	109,
	33,
	267,
	63,
	63,
	33,
	58,
	58,
	50,
	58,
	58,
	63,
	300
];

Excel.run(async (context) =>
{
	let sheet;
	let length = g_sSheetName.length;
	
	if( g_bRunActiveOnly )
	{
		length = 1;
	}
	
	for(let i=0; i<length; i++)
	{
		if( g_bRunActiveOnly )
		{
			sheet = context.workbook.worksheets.getActiveWorksheet();
		}
		else
		{
			sheet = context.workbook.worksheets.getItem(`${g_sSheetName[i]}`);
		}
		sheet.load("name");
		await context.sync();
		
		if(sheet.name !== "Sales Review")
		{
			console.log( `=========== Sheeet "${sheet.name}" Loaded ==========`);
			sheet.activate();
			//sheet.protection.unprotect("1");
			sheet.protection.protect({}, "1");
			continue;
			
			
			//----------------------- --------------------------------------//
			//------------ COLUMN, ROW, SIZE AND TAG STRUCTURE -------------//
			//------------------------ -------------------------------------//
			console.log("*** Building Row, Column, Size and Tag ***");
			sheet.freezePanes.unfreeze();
			BuildSalesStructure(sheet);

			//----------------------- --------------------------------------//
			//--------------------    INSERT FORMULA    --------------------//
			//------------------------ -------------------------------------//
			await context.sync();
			console.log("************ Inserting Formula ***********");
			InsertFormula(sheet);


			//----------------------- --------------------------------------//
			//------------    INSERT CONDITIONAL FORMATTING    -------------//
			//------------------------ -------------------------------------//
			await context.sync();
			console.log("**** Inserting Conditional Formatting ****");
			InsertCondFormatting(sheet);


			//----------------------- --------------------------------------//
			//---------------    HIDE EXTRA ROW & COLUMN    ----------------//
			//------------------------ -------------------------------------//
			await context.sync();
			console.log("******** Hide Extra Row and Column *******");
			HideExtraRowColumn(sheet);


			//----------------------- --------------------------------------//
			//-----------------    ECO REPAIR COMPLETED    -----------------//
			//------------------------ -------------------------------------//
			//sheet.freezePanes.freezeColumns(1);
			sheet.freezePanes.freezeRows(g_iSaleHeader);
			sheet.protection.protect({}, "1");
		}
		await context.sync().then(() => { console.log(`======= "${sheet.name}" repair completed =======`); });
	}
})

function BuildSalesStructure(sheet)
{
	// reset all row height and value
	let range = sheet.getRange();
	if (g_bResetCellValue) {
		range.clear();
	}
	if (g_bResetFillColor) {
		range.format.fill.clear();
	}
	if (g_bResetConFormat) {
		range.conditionalFormats.clearAll();
	}
	//sheet.getRange().rowHidden			= false;
	//sheet.getRange().columnHidden			= false;
	range.format.protection.locked			= true;
	range.format.protection.formulaHidden	= false;
	range.format.font.bold					= false;
	range.format.font.italic				= false;
	range.format.rowHeight					= 20;
	range.format.font.size					= 12;
	range.format.font.name					= "Calibri";
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
	range.unmerge();
	
	ResetLettersBuffer( "", "", "", "A" );
	for(let i=0; i<g_sDescName.length; i++)
	{
		g_sLetterTemps1 = IncreaseLetters(g_sLetterBuff3, g_sLetterBuff2, g_sLetterBuff1, g_sLetterBuff0);
	}
	sheet.getRange("1:" + g_iFootEnd).rowHidden = false;
	sheet.getRange("A:" + g_sLetterTemps1 ).columnHidden = false;
	
	
	// count column length
	ResetLettersBuffer( "", "", "", "A" );
	for(let i=0; i<g_sDescName.length; i++)
	{
		g_sLetterTemps1 = IncreaseLetters(g_sLetterBuff3, g_sLetterBuff2, g_sLetterBuff1, g_sLetterBuff0);
	}
	
	// format background color
	range = sheet.getRange("A1:" + g_sLetterTemps1 + g_iFootEnd);
	range.format.fill.color = g_Color_Default;

	// writeable area color
	range = sheet.getRange("C1:E2");
	range.format.fill.clear();
	
	g_sLetterTemps1 = GetColumnLetters( "STATUS" );
	range = sheet.getRange("A" + g_iBodyStart + ":" + g_sLetterTemps1 + g_iBodyEnd);
	range.format.fill.clear();
	range.format.protection.locked = false;
	range.format.protection.formulaHidden = false;

	g_sLetterTemps1 = GetColumnLetters( "D/DELIVER" );
	g_sLetterTemps2 = GetColumnLetters( "COMMENT" );
	range = sheet.getRange(g_sLetterTemps1 + g_iBodyStart + ":" + g_sLetterTemps2 + g_iBodyEnd);
	range.format.fill.clear();
	range.format.protection.locked = false;
	range.format.protection.formulaHidden = false;
	
	g_sLetterTemps1 = GetColumnLetters( "DRAWING No." );
	range = sheet.getRange(g_sLetterTemps1 + g_iBodyStart + ":" + g_sLetterTemps1 + g_iBodyEnd);
	range.format.horizontalAlignment = "Left";

	g_sLetterTemps1 = GetColumnLetters( "DESCRIPTION" );
	range = sheet.getRange(g_sLetterTemps1 + g_iBodyStart + ":" + g_sLetterTemps1 + g_iBodyEnd);
	range.format.horizontalAlignment = "Left";

	g_sLetterTemps1 = GetColumnLetters( "COMMENT" );
	range = sheet.getRange(g_sLetterTemps1 + g_iBodyStart + ":" + g_sLetterTemps1 + g_iBodyEnd);
	range.format.horizontalAlignment = "Left";

	// sales summary tag
	range = sheet.getRange("B1");
	range.values = "JOB QTY : ";
	range.format.horizontalAlignment = "Right";
	range = sheet.getRange("B2");
	range.values = "COMPLETED : ";
	range.format.horizontalAlignment = "Right";
	range = sheet.getRange("C1:E2");
	range.format.fill.color = g_Color_White;
	range = sheet.getRange("D1:D2");
	range.values = "RM";
	range.format.horizontalAlignment = "Right";
	range = sheet.getRange("E1:E2");
	range.format.horizontalAlignment = "Left";

	// sales description tags
	for(let i=0; i<g_sDescName.length; i++)
	{
		g_sLetterTemps1 = GetColumnLetters( g_sDescName[i] );
		range = sheet.getRange(g_sLetterTemps1 + g_iSaleHeader);
		range.values = g_sDescName[i];
		range.format.columnWidth = g_iDescSize[i];
		range.format.fill.color = g_Color_Black;
		range.format.font.color = g_Color_White;
		range.format.font.bold = true;
	}
}

function InsertFormula(sheet)
{
	sheet.getRange("C1").formulas = `=COUNTA(J${g_iBodyStart}:J${g_iBodyEnd})`;
	sheet.getRange("C2").formulas = `=COUNTIF(M${g_iBodyStart}:M${g_iBodyEnd},"F")`;
	sheet.getRange("E1").formulas = `=SUM(N${g_iBodyStart}:N${g_iBodyEnd})`;
	sheet.getRange("E2").formulas = `=SUM(O${g_iBodyStart}:O${g_iBodyEnd})`;

	for (let i = g_iBodyStart; i <= g_iBodyEnd; i++) {
		let iii = "I" + i;
		let lll = "L" + i;
		sheet.getRange("M" + i).formulas = `=IF(AND(${iii}<>"",${lll}<>${CODE_COMPLETED},${lll}<>${CODE_PENDING},${lll}<>${CODE_FORWARD},${lll}<>${CODE_CANCLED}),${iii}-TODAY(),IF(${lll}=${CODE_COMPLETED},${CODE_COMPLETED},IF(${lll}=${CODE_PENDING},${iii}-TODAY(),IF(${lll}=${CODE_FORWARD},${CODE_FORWARD},IF(${lll}=${CODE_CANCLED},"CANCEL","")))))`;
		sheet.getRange("N" + i).formulas = "=K" + i + "*" + "J" + i;
		sheet.getRange("O" + i).formulas = '=IF(M' + i + '=' + CODE_COMPLETED + ', K' + i + '* J' + i + ', 0)';
	}
}

function InsertCondFormatting(sheet)
{
	let i, range, conditionalFormula;

	for (i = g_iBodyStart; i <= g_iBodyEnd; i++) {
		//=$L$14=COMPLETE
		g_sLetterTemps1 = GetColumnLetters( "COMMENT" );
		g_sLetterTemps2 = GetColumnLetters( "STATUS" );
		range = sheet.getRange("A" + i + ":" + g_sLetterTemps1 + i);
		
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = "=$" + g_sLetterTemps2 + "$" + i + "=" + CODE_COMPLETED;
		conditionalFormula.custom.format.font.color = COLOR_COMPLETED;
		
		//=$L$14=PENDING
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = "=$" + g_sLetterTemps2 + "$" + i + "=" + CODE_PENDING;
		conditionalFormula.custom.format.font.color = COLOR_PENDING;
		
		//=$L$102="HS"
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = `=$${g_sLetterTemps2}$${i}=${CODE_HAVESTOCK}`;
		conditionalFormula.custom.format.font.color = COLOR_HAVESTOCK;
		
		//=$L$14=FORWARD
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = "=$" + g_sLetterTemps2 + "$" + i + "=" + CODE_FORWARD;
		conditionalFormula.custom.format.font.color = COLOR_FORWARD;
		
		//=$L$14=CANCEL
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = "=$" + g_sLetterTemps2 + "$" + i + "=" + CODE_CANCLED;
		conditionalFormula.custom.format.font.color = COLOR_CANCLED;
		
		//=AND($M$14>URGENT,$M$14<=NEARDUE)
		g_sLetterTemps2 = GetColumnLetters( "DAY DUE" );
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = "=AND($" + g_sLetterTemps2 + "$" + i + ">" + CODE_URGENT + ",$" + g_sLetterTemps2 + "$" + i + "<=" + CODE_NEARDUE;
		conditionalFormula.custom.format.font.color = COLOR_NEARDUE;
		
		//=$M$14<=URGENT
		conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		conditionalFormula.custom.rule.formula = "=$" + g_sLetterTemps2 + "$" + i + "<=" + CODE_URGENT;
		conditionalFormula.custom.format.font.color = COLOR_URGENT;
	}
}

function HideExtraRowColumn(sheet)
{
	g_sLetterTemps1 = GetColumnLetters( "ACTUAL SALES" );
	sheet.getRange(g_sLetterTemps1 + "1").columnHidden = true;
}

function IncreaseLetters(charInput3, charInput2, charInput1, charInput0)
{
	let charIn0 = charInput0;
	let charIn1 = charInput1;
	let charIn2 = charInput2;
	let charIn3 = charInput3;
	
	let charIn0 = String.fromCharCode(charIn0.charCodeAt(0) + 1 );
	if( charIn0.charCodeAt(0) > 90 )
	{
		charIn0 = String.fromCharCode( 64 + (charIn0.charCodeAt(0)) - 90 );
		if(charIn1 === "" )
		{
			charIn1 = "A";
		}
		else
		{
			charIn1 = String.fromCharCode( charIn1.charCodeAt(0) + 1 );
			if( charIn1.charCodeAt(0) > 90 )
			{
				charIn1 = String.fromCharCode( 64 + (charIn1.charCodeAt(0)) - 90 );
				if(charIn2 === "" )
				{
					charIn2 = "A";
				}
				else
				{
					charIn2 = String.fromCharCode( charIn1.charCodeAt(0) + 1 );
					if( charIn2.charCodeAt(0) > 90 )
					{
						charIn2 = String.fromCharCode( 64 + (charIn2.charCodeAt(0)) - 90 );
						if(charIn3 === "" )
						{
							charIn3 = "A";
						}
						else
						{
							charIn3 = String.fromCharCode( charIn3.charCodeAt(0) + 1 );
							if( charIn3.charCodeAt(0) > 90 )
							{
								console.log("function 'IncreaseLetters()' has run out of buffers...");
							}
						}
					}
				}
			}
		}
	}
	g_sLetterBuff0 = charIn0;
	g_sLetterBuff1 = charIn1;
	g_sLetterBuff2 = charIn2;
	g_sLetterBuff3 = charIn3;
	
	return ( g_sLetterBuff3 + g_sLetterBuff2 + g_sLetterBuff1 + g_sLetterBuff0 );
}

function ResetLettersBuffer(charInput4, charInput3, charInput2, charInput1)
{
	g_sLetterBuff0 = charInput1;
	g_sLetterBuff1 = charInput2;
	g_sLetterBuff2 = charInput3;
	g_sLetterBuff3 = charInput4;
}

function GetColumnLetters( enddesc )
{
	for(let i=0; i<g_sDescName.length; i++)
	{
		if(enddesc === g_sDescName[i])
		{
			return String.fromCharCode(65 + i);
		}
	}
	
}

