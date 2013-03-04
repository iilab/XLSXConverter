var EXCEL_CONST_SENTINEL = 1;

/* ------ */
function decode_row(rowstr) { return Number(unfix_row(rowstr)) - 1; }
function encode_row(row) { return "" + (row + 1); }
function fix_row(cstr) { return cstr.replace(/([A-Z]|^)([0-9]+)$/,"$1$$$2"); }
function unfix_row(cstr) { return cstr.replace(/\$([0-9]+)$/,"$1"); }

function decode_col(colstr) { var c = unfix_col(colstr), d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
function fix_col(cstr) { return cstr.replace(/^([A-Z])/,"$$$1"); }
function unfix_col(cstr) { return cstr.replace(/^\$([A-Z])/,"$1"); }

function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/,"$1,$2").split(","); }

/* decode_cell assumes that you are passing a valid cell (not a row/col) */
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }
function fix_cell(cstr) { return fix_col(fix_row(cstr)); }
function unfix_cell(cstr) { return unfix_col(unfix_row(cstr)); }

/* ranges can be individual cells -- magic happens here */
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(cs,ce) {
	if(ce === undefined) return encode_range(cs.s, cs.e);
	if(typeof cs !== 'string') cs = encode_cell(cs); if(typeof ce !== 'string') ce = encode_cell(ce);
	return cs == ce ? cs : cs + ":" + ce;
}

function shift_cell(cell, range) {
	if(cell.cRel) cell.c += range.s.c;
	if(cell.rRel) cell.r += range.s.r;
	cell.cRel = cell.rRel = 0;
	return cell;
}

function shift_range(cell, range) {
	cell.s = shift_cell(cell.s, range);
	cell.e = shift_cell(cell.e, range);
	return cell;
}

/* ------ */


/* 2.5.198.2 */
var BERR = {
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
	0xFF: "#WTF?"
};

/* 2.5.198.4 */
var Cetab = {
	0x0000: 'BEEP'
};

/* 2.5.198.17 */
var Ftab = {
	0x0000: 'COUNT',
	0x0001: 'IF',
	0x0004: 'SUM',
	0x017B: 'RTD'
};



/* 2.5.198.44 */
var PtgDataType = {
	0x1: "REFERENCE", // reference to range
	0x2: "VALUE", // single value
	0x3: "ARRAY" // array of values
};

/* 2.5.198.105 RgceArea */
function parse_RgceArea(blob, length) {
	var read = blob.read_shift.bind(blob);
	var r=read(2), R=read(2), c=read(2), C=read(2);
	var cRel = (c >> 14) & 1, rRel = (c >> 15 & 1);
	var CRel = (C >> 14) & 1, RRel = (C >> 15 & 1);
	c &= 0xFF; C &= 0xFF;
	return {s:{r:r,c:c,cRel:cRel, rRel:rRel},e:{r:R,c:C,cRel:CRel,rRel:RRel}};
}

/* 2.5.198.27 TODO */
function parse_PtgArea(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var area = parse_RgceArea(blob, 8);
	return [type, area];
}

/* 2.5.198.109 */
function parse_RgceLoc(blob, length) {
	var rw = blob.read_shift(2);
	var cl = blob.read_shift(2);
	var cRel = cl & 0x80, rRel = cl & 0x40;
	cl &= 0x3F; 
	return {r:rw,c:cl,cRel:cRel,rRel:rRel};
}

/* 2.5.198.84 TODO */
function parse_PtgRef(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLoc(blob,4);
	return [type, loc];
}

/* 2.5.198.84 TODO */
function parse_PtgRef3d(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var ixti = blob.read_shift(2); // XtiIndex
	var loc = parse_RgceLoc(blob,4);
	return [type, ixti, loc];
}


function parseread5(blob, length) { blob.l+=5; return; }
function parseread4(blob, length) { blob.l+=4; return; }
function parseread1(blob, length) { blob.l+=1; return; }

/* 2.5.198.35 TODO */
function parse_PtgAttrGoto(blob, length) {
	blob.l += 2;
	return blob.read_shift(2);
}


/* 2.5.198.63 TODO */
function parse_PtgFuncVar(blob, length) {
	blob.l++; 
	var cparams = blob.read_shift(1), tab = parsetab(blob);
	return [cparams, (tab[0] === 0 ? Ftab : Cetab)[tab[1]]];
}

function parsetab(blob, length) {
	return [blob[blob.l+1]>>7, blob.read_shift(2) & 0x7FFF];
}


/* 2.5.198.36 */
var parse_PtgAttrIf = parseread4;
/* 2.5.198.41 */
var parse_PtgAttrSum = parseread4; 

/* 2.5.198.26 */
var parse_PtgAdd = parseread1;
/* 2.5.198.43 */
var parse_PtgConcat = parseread1;
/* 2.5.198.45 */
var parse_PtgDiv = parseread1;
/* 2.5.198.56 */
var parse_PtgEq = parseread1;
/* 2.5.198.58 TODO */
var parse_PtgExp = parseread5;
/* 2.5.198.64 */
var parse_PtgGe = parseread1;
/* 2.5.198.65 */
var parse_PtgGt = parseread1;
/* 2.5.198.66 TODO */
var parse_PtgInt = function(blob, length){blob.l++; return blob.read_shift(2);};
/* 2.5.198.68 */
var parse_PtgLe = parseread1;
/* 2.5.198.69 */
var parse_PtgLt = parseread1;
/* 2.5.198.75 */
var parse_PtgMul = parseread1;
/* 2.5.198.80 */
var parse_PtgParen = parseread1;
/* 2.5.198.82 */
var parse_PtgPower = parseread1;
/* 2.5.198.90 */
var parse_PtgSub = parseread1;

/* 2.5.198.25 */
var PtgTypes = {
	0x01: { n:'PtgExp', f:parse_PtgExp },
	0x03: { n:'PtgAdd', f:parse_PtgAdd },
	0x04: { n:'PtgSub', f:parse_PtgSub },
	0x05: { n:'PtgMul', f:parse_PtgMul },
	0x06: { n:'PtgDiv', f:parse_PtgDiv },
	0x07: { n:'PtgPower', f:parse_PtgPower },
	0x08: { n:'PtgConcat', f:parse_PtgConcat },
	0x09: { n:'PtgLt', f:parse_PtgLt },
	0x0A: { n:'PtgLe', f:parse_PtgLe },
	0x0B: { n:'PtgEq', f:parse_PtgEq },
	0x0C: { n:'PtgGe', f:parse_PtgGe },
	0x0D: { n:'PtgGt', f:parse_PtgGt },
	0x15: { n:'PtgParen', f:parse_PtgParen },
	0x1E: { n:'PtgInt', f:parse_PtgInt },
	0x22: { n:'PtgFuncVar', f:parse_PtgFuncVar },
	0x24: { n:'PtgRef', f:parse_PtgRef },
	0x25: { n:'PtgArea', f:parse_PtgArea },
	0x42: { n:'PtgFuncVar', f:parse_PtgFuncVar },
	0x44: { n:'PtgRef', f:parse_PtgRef },
	0x5A: { n:'PtgRef3d', f:parse_PtgRef3d },

	0xFF: {}
};

var Ptg18 = {};
var Ptg19 = {
	0x02: { n:'PtgAttrIf', f:parse_PtgAttrIf },
	0x08: { n:'PtgAttrGoto', f:parse_PtgAttrGoto },
	0x10: { n:'PtgAttrSum', f:parse_PtgAttrSum },
	0xFF: {}
};

/* sections refer to MS-XLS unless otherwise stated */

/* --- Simple Utilities --- */
function parsenoop(blob, length) { blob.read_shift(length); return; }
function parsenoop2(blob, length) { blob.read_shift(length); return null; }

function parslurp(blob, length, cb) {
	var arr = [], target = blob.l + length;
	while(blob.l < target) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw "Slurp error";
	return arr;
}

function parslurp2(blob, length, cb) {
	var arr = [], target = blob.l + length, len = blob.read_shift(2);
	while(len-- !== 0) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw "Slurp error";
	return arr;
}

function parseuint16(blob, length) { return blob.read_shift(2, 'u'); } 
function parseuint16a(blob, length) { return parslurp(blob,length,parseuint16);}



/* --- 2.5 Structures --- */

/* 2.5.14 */
function parsebool(blob, length) { return blob.read_shift(length) === 0x1; }
var parse_Boolean = parsebool;

/* 2.5.19 */
function parse_Cell(blob, length) {
	var rw = blob.read_shift(2); // 0-indexed
	var col = blob.read_shift(2);
	var ixfe = blob.read_shift(2);
	return {r:rw, c:col, ixfe:ixfe};
}

/* 2.5.134 */
function parse_frtHeader(blob) {
	var read = blob.read_shift.bind(blob);
	var rt = read(2);
	var flags = read(2); // TODO: parse these flags
	read(8);
	return {type: rt, flags: flags};
}

/* 2.5.240 */
function parse_ShortXLUnicodeString(blob) {
	var read = blob.read_shift.bind(blob);
	var cch = read(1);
	var fHighByte = read(1);
	var retval;
	if(fHighByte===0) { retval = blob.utf8(blob.l, blob.l+cch); blob.l += cch; }
	else { retval = blob.utf16le(blob.l, blob.l + 2*cch); blob.l += 2*cch; }
	return retval;
}

/* 2.5.293 */
function parse_XLUnicodeRichExtendedString(blob) {
	var cch = blob.read_shift(2), flags = blob.read_shift(1);
	var width = 1 + (flags & 0x1);
	// TODO: cRun
	// TODO: cbExtRst
	var msg = blob.read_shift('utf8', cch);
	return msg;
}

/* 2.5.294 */
function parse_XLUnicodeString(blob) {
	var read = blob.read_shift.bind(blob);
	var cch = read(2);
	var fHighByte = read(1);
	var retval;
	if(fHighByte===0) { retval = blob.utf8(blob.l, blob.l+cch); blob.l += cch; }
	else { retval = blob.utf16le(blob.l, blob.l + 2*cch); blob.l += 2*cch; }
	return retval;
}

function parse_OptXLUnicodeString(blob, length) { return length === 0 ? "" : parseXLUnicodeString(blob); }

/* 2.5.342 */
function parse_Xnum(blob, length) { return blob.read_shift('ieee754'); }

/* 2.5.158 */
var HIDEOBJENUM = ['SHOWALL', 'SHOWPLACEHOLDER', 'HIDEALL'];
var parse_HideObjEnum = parseuint16; 

function parse_XTI(blob, length) {
	var read = blob.read_shift.bind(blob);
	var iSupBook = read(2), itabFirst = read(2,'i'), itabLast = read(2,'i');
	return [iSupBook, itabFirst, itabLast];
}
function parse_XTI2(blob, length) { return parslurp2(blob,length,parse_XTI);}

/* 2.5.217 */
function parse_RkNumber(blob) {
	var b = blob.slice(blob.l, blob.l+4);
	var div100 = b[0] & 1, fInt = b[0] & 2;
	blob.l+=4;
	b[0] &= ~3;
	var RK = [0,0,0,0,b[0],b[1],b[2],b[3]].readDoubleLE(0);
	// 30 most significant bits ..
	return div100 ? RK/100 : RK;
}

/* 2.5.218 */
function parse_RkRec(blob, length) {
	var ixfe = blob.read_shift(2);
	var RK = parse_RkNumber(blob);
	return [ixfe, RK];
}


/* 2.5.133 */
function parse_FormulaValue(blob) {
	var b;
	if(blob.readUInt16LE(blob.l + 6) !== 0xFFFF) return parse_Xnum(blob);
	switch(blob[blob.l+2]) {
		case 0x00: blob.l += 8; return "String";
		case 0x01: b = blob[blob.l+2] === 0x1; blob.l += 8; return b; 
		case 0x02: b = BERR[blob.l+2]; blob.l += 8; return b;
		case 0x03: blob.l += 8; return "";
	}
}

/* 2.5.198.104 */
var parse_Rgce = function(blob, length) {
	var target = blob.l + length;
	var R, id, ptgs = [];
	while(target != blob.l) {
		length = target - blob.l;
		id = blob[blob.l];
		R = PtgTypes[id];
		if(id === 0x18 || id === 0x19) {
			id = blob[blob.l + 1];
			R = (id === 0x18 ? Ptg18 : Ptg19)[id];
		}
		if(!R) { ptgs.push(parsenoop(blob, length)); }
		else { ptgs.push([R.n, R.f(blob, length)]); }
	}
	return ptgs;
};

/* */
var parse_RgbExtra = parsenoop;

/* 2.5.198.3 TODO */
function parse_CellParsedFormula(blob, length) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	var rgce = parse_Rgce(blob, cce);
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce);
	return [rgce, rgcb];
}

/* --- 2.4 Records --- */

/* 2.4.21 */
function parse_BOF(blob, length) {
	var o = {};
	o.BIFFVer = blob.read_shift(2); length -= 2;
	if(o.BIFFVer != 0x0600) throw "Unexpected BIFF Ver " + o.BIFFVer;
	blob.read_shift(length);
	return o;	
}


/* 2.4.146 */
function parse_InterfaceHdr(blob, length) {
	if((q=blob.read_shift(2))!==0x04b0) throw 'InterfaceHdr codePage ' + q;
	return 0x04b0;
}


/* 2.4.349 */
function parse_WriteAccess(blob, length) {
	var l = blob.l;
	var UserName = parse_XLUnicodeString(blob);
	blob.read_shift(length + l - blob.l);
	return { WriteAccess: UserName };
}

/* 2.4.28 */
function parse_BoundSheet8(blob, length) {
	var read = blob.read_shift.bind(blob);
	var pos = read(4); 
	var hidden = read(1) >> 6;
	var dt = read(1);
	switch(dt) {
		case 0: dt = 'Worksheet'; break;
		case 1: dt = 'Macrosheet'; break;
		case 2: dt = 'Chartsheet'; break;
		case 6: dt = 'VBAModule'; break;
	}
	var name = parse_ShortXLUnicodeString(blob);
	return { pos:pos, hs:hidden, dt:dt, name:name };
}

/* 2.4.265 TODO */
function parse_SST(blob, length) {
	if(length > 1024) return [parsenoop(blob,length)];
	var read = blob.read_shift.bind(blob);
	var cnt = read(4);
	var ucnt = read(4);
	var strs = [];
	for(var i = 0; i != ucnt; ++i) {
		strs.push(parse_XLUnicodeRichExtendedString(blob));
	}
	strs.Count = cnt; strs.Unique = ucnt;
	return strs;
	//return { Count:cnt, Unique:ucnt, Strings:strs};	
}


/* 2.4.221 TODO*/
function parse_Row(blob, length) {
	var read = blob.read_shift.bind(blob);
	var rw = read(2), col = read(2), Col = read(2), rht = read(2);
	read(4); // reserved(2), unused(2)
	var flags = read(1); // various flags
	read(1); // reserved
	read(2); //ixfe, other flags
	return {r:rw, c:col, cnt:Col-col};
}


/* 2.4.125 */
function parse_ForceFullCalculation(blob, length) {
	var header = parse_frtHeader(blob);
	if(header.type != 0x08A3) throw "Invalid Future Record " + header.type;
	var fullcalc = blob.read_shift(4);
	return { FullCalc: fullcalc };
}


var parse_CompressPictures = parsenoop2; /* 2.4.55 Not interesting */



/* 2.4.215 rt */
function parse_RecalcId(blob, length) {
	blob.read_shift(2);
	return blob.read_shift(4);
}

/* 2.4.87 */
function parse_DefaultRowHeight (blob, length) {
	var f = blob.read_shift(2), miyRw;
	miyRw = blob.read_shift(2); // flags & 0x02 -> hidden, else empty
	var fl = {Unsynced:f&1,DyZero:(f&2)>>1,ExAsc:(f&4)>>2,ExDsc:(f&8)>>3};
	return [fl, miyRw];
}

/* 2.4.345 TODO */
function parse_Window1(blob, length) {
	var read = blob.read_shift.bind(blob);
	var xWn = read(2), yWn = read(2), dxWn = read(2), dyWn = read(2);
	var flags = read(2), iTabCur = read(2), iTabFirst = read(2);
	var ctabSel = read(2), wTabRatio = read(2);
	return { Pos: [xWn, yWn], Dim: [dxWn, dyWn], Flags: flags, CurTab: iTabCur,
		FirstTab: iTabFirst, Selected: ctabSel, TabRatio: wTabRatio };
}

/* 2.4.122 TODO */
function parse_Font(blob, length) {
	blob.l += 14;
	var name = parse_ShortXLUnicodeString(blob);
	return name;
}

/* 2.4.149 */
function parse_LabelSst(blob, length) {
	var cell = parse_Cell(blob);
	cell.isst = blob.read_shift(4);
	return cell;
}

/* 2.4.126 TODO: ABNF */
function parse_Format(blob, length) {
	var ifmt = blob.read_shift(2);
	return [ifmt, parse_XLUnicodeString(blob)];
}

/* 2.4.90 */
function parse_Dimensions(blob, length) {
	var read = blob.read_shift.bind(blob);
	var r = read(4), R = read(4), c = read(2), C = read(2);
	read(2);
	return {s: {r:r, c:c}, e: {r:R, c:C}};
}

/* 2.4.220 */
function parse_RK(blob, length) {
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var rkrec = parse_RkRec(blob);
	return {r:rw, c:col, ixfe:rkrec[0], rknum:rkrec[1]};
}

/* 2.4.175 */
function parse_MulRk(blob, length) {
	var target = blob.l + length - 2;
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var rkrecs = [];
	while(blob.l < target) rkrecs.push(parse_RkRec(blob));
	if(blob.l !== target) throw "MulRK read error";
	var lastcol = blob.read_shift(2);
	if(rkrecs.length != lastcol - col + 1) throw "MulRK length mismatch";
	return {r:rw, c:col, C:lastcol, rkrec:rkrecs};
}


/* 2.4.127 TODO*/ 
function parse_Formula(blob, length) {
	var cell = parse_Cell(blob, 6);
	var val = parse_FormulaValue(blob,8);
	var flags = blob.read_shift(1);
	blob.read_shift(1);
	var chn = blob.read_shift(4);
	var cbf = parse_CellParsedFormula(blob, length-20);
	return {cell:cell, val:val, formula:cbf};
}

function parse_Number(blob, length) {
	var cell = parse_Cell(blob, 6);
	var xnum = parse_Xnum(blob, 8);
	cell.val = xnum;
	return cell;
}

var parse_XLHeaderFooter = parse_OptXLUnicodeString; // TODO: parse 2.4.136


var parse_Backup = parsebool; /* 2.4.14 */
var parse_Blank = parse_Cell; /* 2.4.20 Just the cell */ 
var parse_BottomMargin = parse_Xnum; /* 2.4.27 */
var parse_BuiltInFnGroupCount = parseuint16; /* 2.4.30 0x0E or 0x10*/ 
var parse_CalcCount = parseuint16; /* 2.4.31 #Iterations */
var parse_CalcDelta = parse_Xnum; /* 2.4.32 */ 
var parse_CalcIter = parsebool;  /* 2.4.33 1=iterative calc */
var parse_CalcMode = parseuint16; /* 2.4.34 0=manual, 1=auto (def), 2=table */
var parse_CalcPrecision = parsebool; /* 2.4.35 */
var parse_CalcRefMode = parsenoop2; /* 2.4.36 */
var parse_CalcSaveRecalc = parsebool; /* 2.4.37 */
var parse_CodePage = parseuint16; /* 2.4.52 */
var parse_Compat12 = parsebool; /* 2.4.54 true = no compatibility check */ 
var parse_Country = parseuint16a; /* 2.4.63 -- two ints, 1 to 981 */ 
var parse_Date1904 = parsebool; /* 2.4.77 - 1=1904,0=1900 */ 
var parse_DefColWidth = parseuint16; /* 2.4.89 */ 
var parse_DSF = parsenoop2; /* 2.4.94 -- MUST be ignored */ 
var parse_EntExU2 = parsenoop2; /* 2.4.102 -- Explicitly says to ignore */
var parse_EOF = parsenoop2; /* 2.4.103 */
var parse_Excel9File = parsenoop2; /* 2.4.104 -- Optional and unused */ 
var parse_ExternSheet = parse_XTI2; /* 2.4.106 */ 
var parse_FeatHdr = parsenoop2; /* 2.4.112 */
var parse_FontX = parseuint16; /* 2.4.123 */ 
var parse_Footer = parse_XLHeaderFooter; /* 2.4.124 */ 
var parse_GridSet = parseuint16; /* 2.4.132, =1 */ 
var parse_HCenter = parsebool; /* 2.4.135 sheet centered horizontal on print */
var parse_Header = parse_XLHeaderFooter; /* 2.4.136 */ 
var parse_HideObj = parse_HideObjEnum; /* 2.4.139 */ 
var parse_InterfaceEnd = parsenoop2; /* 2.4.145 -- noop */
var parse_LeftMargin = parse_Xnum; /* 2.4.151 */ 
var parse_Mms = parsenoop2; /* 2.4.169 */
var parse_ObjProtect = parsebool; /* 2.4.183 -- must be 1 if present */ 
var parse_Password = parseuint16; /* 2.4.191 */
var parse_PrintGrid = parsebool; /* 2.4.202 */ 
var parse_PrintRowCol = parsebool; /* 2.4.203 */ 
var parse_PrintSize = parseuint16; /* 2.4.204 0:3 */ 
var parse_Prot4Rev = parsebool; /* 2.4.205 */
var parse_Prot4RevPass = parseuint16; /* 2.4.206 */
var parse_Protect = parsebool; /* 2.4.207 */ 
var parse_RefreshAll = parsebool; /* 2.4.217 -- must be 0 if not template */ 
var parse_RightMargin = parse_Xnum; /* 2.4.219 */ 
var parse_RRTabId = parseuint16a; /* 2.4.241 */
var parse_ScenarioProtect = parsebool; /* 2.4.245 */ 
var parse_Scl = parseuint16a; /* 2.4.247 num, den */ 
var parse_String = parse_XLUnicodeString; /* 2.4.268 */
var parse_SxBool = parsebool; /* 2.4.274 */ 
var parse_TopMargin = parse_Xnum; /* 2.4.328 */ 
var parse_UsesELFs = parsebool; /* 2.4.337 -- should be 0 */
var parse_VCenter = parsebool; /* 2.4.342 */ 
var parse_WinProtect = parsebool; /* 2.4.347 */


/* ---- */
var parse_Lbl = parsenoop; 
var parse_VerticalPageBreaks = parsenoop; 
var parse_HorizontalPageBreaks = parsenoop; 
var parse_Note = parsenoop; 
var parse_Selection = parsenoop; 
var parse_ExternName = parsenoop; 
var parse_FilePass = parsenoop; 
var parse_Continue = parsenoop; 
var parse_Pane = parsenoop; 
var parse_Pls = parsenoop; 
var parse_DCon = parsenoop; 
var parse_DConRef = parsenoop; 
var parse_DConName = parsenoop; 
var parse_XCT = parsenoop; 
var parse_CRN = parsenoop; 
var parse_FileSharing = parsenoop; 
var parse_Obj = parsenoop; 
var parse_Uncalced = parsenoop; 
var parse_Template = parsenoop; 
var parse_Intl = parsenoop; 
var parse_ColInfo = parsenoop; 
var parse_Guts = parsenoop; 
var parse_WsBool = parsenoop; 
var parse_WriteProtect = parsenoop; 
var parse_Sort = parsenoop; 
var parse_Palette = parsenoop; 
var parse_Sync = parsenoop; 
var parse_LPr = parsenoop; 
var parse_DxGCol = parsenoop; 
var parse_FnGroupName = parsenoop; 
var parse_FilterMode = parsenoop; 
var parse_AutoFilterInfo = parsenoop; 
var parse_AutoFilter = parsenoop; 
var parse_Setup = parsenoop; 
var parse_ScenMan = parsenoop; 
var parse_SCENARIO = parsenoop; 
var parse_SxView = parsenoop; 
var parse_Sxvd = parsenoop; 
var parse_SXVI = parsenoop; 
var parse_SxIvd = parsenoop; 
var parse_SXLI = parsenoop; 
var parse_SXPI = parsenoop; 
var parse_DocRoute = parsenoop; 
var parse_RecipName = parsenoop; 
var parse_MulBlank = parsenoop; 
var parse_SXDI = parsenoop; 
var parse_SXDB = parsenoop; 
var parse_SXFDB = parsenoop; 
var parse_SXDBB = parsenoop; 
var parse_SXNum = parsenoop; 
var parse_SxErr = parsenoop; 
var parse_SXInt = parsenoop; 
var parse_SXString = parsenoop; 
var parse_SXDtr = parsenoop; 
var parse_SxNil = parsenoop; 
var parse_SXTbl = parsenoop; 
var parse_SXTBRGIITM = parsenoop; 
var parse_SxTbpg = parsenoop; 
var parse_ObProj = parsenoop; 
var parse_SXStreamID = parsenoop; 
var parse_DBCell = parsenoop; 
var parse_SXRng = parsenoop; 
var parse_SxIsxoper = parsenoop; 
var parse_BookBool = parsenoop; 
var parse_DbOrParamQry = parsenoop; 
var parse_OleObjectSize = parsenoop; 
var parse_XF = parsenoop; 
var parse_SXVS = parsenoop; 
var parse_MergeCells = parsenoop; 
var parse_BkHim = parsenoop; 
var parse_MsoDrawingGroup = parsenoop; 
var parse_MsoDrawing = parsenoop; 
var parse_MsoDrawingSelection = parsenoop; 
var parse_PhoneticInfo = parsenoop; 
var parse_SxRule = parsenoop; 
var parse_SXEx = parsenoop; 
var parse_SxFilt = parsenoop; 
var parse_SxDXF = parsenoop; 
var parse_SxItm = parsenoop; 
var parse_SxName = parsenoop; 
var parse_SxSelect = parsenoop; 
var parse_SXPair = parsenoop; 
var parse_SxFmla = parsenoop; 
var parse_SxFormat = parsenoop; 
var parse_ExtSST = parsenoop; 
var parse_SXVDEx = parsenoop; 
var parse_SXFormula = parsenoop; 
var parse_SXDBEx = parsenoop; 
var parse_RRDInsDel = parsenoop; 
var parse_RRDHead = parsenoop; 
var parse_RRDChgCell = parsenoop; 
var parse_RRDRenSheet = parsenoop; 
var parse_RRSort = parsenoop; 
var parse_RRDMove = parsenoop; 
var parse_RRFormat = parsenoop; 
var parse_RRAutoFmt = parsenoop; 
var parse_RRInsertSh = parsenoop; 
var parse_RRDMoveBegin = parsenoop; 
var parse_RRDMoveEnd = parsenoop; 
var parse_RRDInsDelBegin = parsenoop; 
var parse_RRDInsDelEnd = parsenoop; 
var parse_RRDConflict = parsenoop; 
var parse_RRDDefName = parsenoop; 
var parse_RRDRstEtxp = parsenoop; 
var parse_LRng = parsenoop; 
var parse_CUsr = parsenoop; 
var parse_CbUsr = parsenoop; 
var parse_UsrInfo = parsenoop; 
var parse_UsrExcl = parsenoop; 
var parse_FileLock = parsenoop; 
var parse_RRDInfo = parsenoop; 
var parse_BCUsrs = parsenoop; 
var parse_UsrChk = parsenoop; 
var parse_UserBView = parsenoop; 
var parse_UserSViewBegin = parsenoop; // overloaded 
var parse_UserSViewEnd = parsenoop; 
var parse_RRDUserView = parsenoop; 
var parse_Qsi = parsenoop; 
var parse_SupBook = parsenoop; 
var parse_CondFmt = parsenoop; 
var parse_CF = parsenoop; 
var parse_DVal = parsenoop; 
var parse_DConBin = parsenoop; 
var parse_TxO = parsenoop; 
var parse_HLink = parsenoop; 
var parse_Lel = parsenoop; 
var parse_CodeName = parsenoop; 
var parse_SXFDBType = parsenoop; 
var parse_ObNoMacros = parsenoop; 
var parse_Dv = parsenoop; 
var parse_Label = parsenoop; 
var parse_BoolErr = parsenoop; 
var parse_Index = parsenoop; 
var parse_Array = parsenoop; 
var parse_Table = parsenoop; 
var parse_Window2 = parsenoop; 
var parse_Style = parsenoop; 
var parse_BigName = parsenoop; 
var parse_ContinueBigName = parsenoop; 
var parse_ShrFmla = parsenoop; 
var parse_HLinkTooltip = parsenoop; 
var parse_WebPub = parsenoop; 
var parse_QsiSXTag = parsenoop; 
var parse_DBQueryExt = parsenoop; 
var parse_ExtString = parsenoop; 
var parse_TxtQry = parsenoop; 
var parse_Qsir = parsenoop; 
var parse_Qsif = parsenoop; 
var parse_RRDTQSIF = parsenoop; 
var parse_OleDbConn = parsenoop; 
var parse_WOpt = parsenoop; 
var parse_SXViewEx = parsenoop; 
var parse_SXTH = parsenoop; 
var parse_SXPIEx = parsenoop; 
var parse_SXVDTEx = parsenoop; 
var parse_SXViewEx9 = parsenoop; 
var parse_ContinueFrt = parsenoop; 
var parse_RealTimeData = parsenoop; 
var parse_ChartFrtInfo = parsenoop; 
var parse_FrtWrapper = parsenoop; 
var parse_StartBlock = parsenoop; 
var parse_EndBlock = parsenoop; 
var parse_StartObject = parsenoop; 
var parse_EndObject = parsenoop; 
var parse_CatLab = parsenoop; 
var parse_YMult = parsenoop; 
var parse_SXViewLink = parsenoop; 
var parse_PivotChartBits = parsenoop; 
var parse_FrtFontList = parsenoop; 
var parse_SheetExt = parsenoop; 
var parse_BookExt = parsenoop; 
var parse_SXAddl = parsenoop; 
var parse_CrErr = parsenoop; 
var parse_HFPicture = parsenoop; 
var parse_Feat = parsenoop; 
var parse_DataLabExt = parsenoop; 
var parse_DataLabExtContents = parsenoop; 
var parse_CellWatch = parsenoop; 
var parse_FeatHdr11 = parsenoop; 
var parse_Feature11 = parsenoop; 
var parse_DropDownObjIds = parsenoop; 
var parse_ContinueFrt11 = parsenoop; 
var parse_DConn = parsenoop; 
var parse_List12 = parsenoop; 
var parse_Feature12 = parsenoop; 
var parse_CondFmt12 = parsenoop; 
var parse_CF12 = parsenoop; 
var parse_CFEx = parsenoop; 
var parse_XFCRC = parsenoop; 
var parse_XFExt = parsenoop; 
var parse_AutoFilter12 = parsenoop; 
var parse_ContinueFrt12 = parsenoop; 
var parse_MDTInfo = parsenoop; 
var parse_MDXStr = parsenoop; 
var parse_MDXTuple = parsenoop; 
var parse_MDXSet = parsenoop; 
var parse_MDXProp = parsenoop; 
var parse_MDXKPI = parsenoop; 
var parse_MDB = parsenoop; 
var parse_PLV = parsenoop; 
var parse_DXF = parsenoop; 
var parse_TableStyles = parsenoop; 
var parse_TableStyle = parsenoop; 
var parse_TableStyleElement = parsenoop; 
var parse_StyleExt = parsenoop; 
var parse_NamePublish = parsenoop; 
var parse_NameCmt = parsenoop; 
var parse_SortData = parsenoop; 
var parse_Theme = parsenoop; 
var parse_GUIDTypeLib = parsenoop; 
var parse_FnGrp12 = parsenoop; 
var parse_NameFnGrp12 = parsenoop; 
var parse_MTRSettings = parsenoop; 
var parse_HeaderFooter = parsenoop; 
var parse_CrtLayout12 = parsenoop; 
var parse_CrtMlFrt = parsenoop; 
var parse_CrtMlFrtContinue = parsenoop; 
var parse_ShapePropsStream = parsenoop; 
var parse_TextPropsStream = parsenoop; 
var parse_RichTextStream = parsenoop; 
var parse_CrtLayout12A = parsenoop; 
var parse_Units = parsenoop; 
var parse_Chart = parsenoop; 
var parse_Series = parsenoop; 
var parse_DataFormat = parsenoop; 
var parse_LineFormat = parsenoop; 
var parse_MarkerFormat = parsenoop; 
var parse_AreaFormat = parsenoop; 
var parse_PieFormat = parsenoop; 
var parse_AttachedLabel = parsenoop; 
var parse_SeriesText = parsenoop; 
var parse_ChartFormat = parsenoop; 
var parse_Legend = parsenoop; 
var parse_SeriesList = parsenoop; 
var parse_Bar = parsenoop; 
var parse_Line = parsenoop; 
var parse_Pie = parsenoop; 
var parse_Area = parsenoop; 
var parse_Scatter = parsenoop; 
var parse_CrtLine = parsenoop; 
var parse_Axis = parsenoop; 
var parse_Tick = parsenoop; 
var parse_ValueRange = parsenoop; 
var parse_CatSerRange = parsenoop; 
var parse_AxisLine = parsenoop; 
var parse_CrtLink = parsenoop; 
var parse_DefaultText = parsenoop; 
var parse_Text = parsenoop; 
var parse_ObjectLink = parsenoop; 
var parse_Frame = parsenoop; 
var parse_Begin = parsenoop; 
var parse_End = parsenoop; 
var parse_PlotArea = parsenoop; 
var parse_Chart3d = parsenoop; 
var parse_PicF = parsenoop; 
var parse_DropBar = parsenoop; 
var parse_Radar = parsenoop; 
var parse_Surf = parsenoop; 
var parse_RadarArea = parsenoop; 
var parse_AxisParent = parsenoop; 
var parse_LegendException = parsenoop; 
var parse_ShtProps = parsenoop; 
var parse_SerToCrt = parsenoop; 
var parse_AxesUsed = parsenoop; 
var parse_SBaseRef = parsenoop; 
var parse_SerParent = parsenoop; 
var parse_SerAuxTrend = parsenoop; 
var parse_IFmtRecord = parsenoop; 
var parse_Pos = parsenoop; 
var parse_AlRuns = parsenoop; 
var parse_BRAI = parsenoop; 
var parse_SerAuxErrBar = parsenoop; 
var parse_ClrtClient = parsenoop; 
var parse_SerFmt = parsenoop; 
var parse_Chart3DBarShape = parsenoop; 
var parse_Fbi = parsenoop; 
var parse_BopPop = parsenoop; 
var parse_AxcExt = parsenoop; 
var parse_Dat = parsenoop; 
var parse_PlotGrowth = parsenoop; 
var parse_SIIndex = parsenoop; 
var parse_GelFrame = parsenoop; 
var parse_BopPopCustom = parsenoop; 
var parse_Fbi2 = parsenoop; 


var RecordEnum = {
	0x0809: { n:'BOF', f:parse_BOF },
	0x000a: { n:'EOF', f:parse_EOF }, 
	
	0x0006: { n:"Formula", f:parse_Formula },
	0x000c: { n:"CalcCount", f:parse_CalcCount },
	0x000d: { n:"CalcMode", f:parse_CalcMode },
	0x000e: { n:"CalcPrecision", f:parse_CalcPrecision },
	0x000f: { n:"CalcRefMode", f:parse_CalcRefMode },
	0x0010: { n:"CalcDelta", f:parse_CalcDelta },
	0x0011: { n:"CalcIter", f:parse_CalcIter },
	0x0012: { n:"Protect", f:parse_Protect },
	0x0013: { n:"Password", f:parse_Password },
	0x0014: { n:"Header", f:parse_Header },
	0x0015: { n:"Footer", f:parse_Footer },
	0x0017: { n:"ExternSheet", f:parse_ExternSheet },
	0x0019: { n:"WinProtect", f:parse_WinProtect },
	0x0022: { n:"Date1904", f:parse_Date1904 },
	0x0028: { n:"TopMargin", f:parse_TopMargin },
	0x0029: { n:"BottomMargin", f:parse_BottomMargin },
	0x0026: { n:"LeftMargin", f:parse_LeftMargin },
	0x0027: { n:"RightMargin", f:parse_RightMargin },
	0x002a: { n:"PrintRowCol", f:parse_PrintRowCol },
	0x002b: { n:"PrintGrid", f:parse_PrintGrid },
	0x0031: { n:"Font", f:parse_Font },
	0x0033: { n:"PrintSize", f:parse_PrintSize },
	0x003d: { n:"Window1", f:parse_Window1 },
	0x0040: { n:"Backup", f:parse_Backup },
	0x0042: { n:'CodePage', f:parse_CodePage },
	0x0055: { n:"DefColWidth", f:parse_DefColWidth },
	0x005c: { n:'WriteAccess', f:parse_WriteAccess },
	0x005f: { n:"CalcSaveRecalc", f:parse_CalcSaveRecalc },
	0x0063: { n:"ObjProtect", f:parse_ObjProtect },
	0x0082: { n:"GridSet", f:parse_GridSet },
	0x0083: { n:"HCenter", f:parse_HCenter },
	0x0084: { n:"VCenter", f:parse_VCenter },
	0x0085: { n:'BoundSheet8', f:parse_BoundSheet8 },
	0x008c: { n:"Country", f:parse_Country },
	0x008d: { n:"HideObj", f:parse_HideObj },
	0x009c: { n:"BuiltInFnGroupCount", f:parse_BuiltInFnGroupCount },
	0x00bd: { n:"MulRk", f:parse_MulRk },
	0x00c1: { n:'Mms', f:parse_Mms },
	0x00ca: { n:"SxBool", f:parse_SxBool },
	0x00dd: { n:"ScenarioProtect", f:parse_ScenarioProtect },
	0x00e1: { n:'InterfaceHdr', f:parse_InterfaceHdr },
	0x00e2: { n:'InterfaceEnd', f:parse_InterfaceEnd },
	0x00fc: { n:"SST", f:parse_SST },
	0x00fd: { n:"LabelSst", f:parse_LabelSst },
	0x013d: { n:"RRTabId", f:parse_RRTabId },
	0x0160: { n:"UsesELFs", f:parse_UsesELFs },
	0x0161: { n:"DSF", f:parse_DSF },
	0x01af: { n:"Prot4Rev", f:parse_Prot4Rev },
	0x01b7: { n:"RefreshAll", f:parse_RefreshAll },
	0x01bc: { n:"Prot4RevPass", f:parse_Prot4RevPass },
	0x01c0: { n:"Excel9File", f:parse_Excel9File },
	0x01c1: { n:"RecalcId", f:parse_RecalcId, r:2},
	0x01c2: { n:"EntExU2", f:parse_EntExU2 },
	0x0200: { n:"Dimensions", f:parse_Dimensions },
	0x0207: { n:"String", f:parse_String },
	0x0208: { n:'Row', f:parse_Row },
	0x0225: { n:"DefaultRowHeight", f:parse_DefaultRowHeight },
	0x041e: { n:"Format", f:parse_Format },
	0x0867: { n:'FeatHdr', f:parse_FeatHdr },
	0x089b: { n:"CompressPictures", f:parse_CompressPictures },
	0x08a3: { n:"ForceFullCalculation", f:parse_ForceFullCalculation },
	0x1026: { n:"FontX", f:parse_FontX },
	
	
	0x0018: { n:"Lbl", f:parse_Lbl },
	0x001a: { n:"VerticalPageBreaks", f:parse_VerticalPageBreaks },
	0x001b: { n:"HorizontalPageBreaks", f:parse_HorizontalPageBreaks },
	0x001c: { n:"Note", f:parse_Note },
	0x001d: { n:"Selection", f:parse_Selection },
	0x0023: { n:"ExternName", f:parse_ExternName },
	0x002f: { n:"FilePass", f:parse_FilePass },
	0x003c: { n:"Continue", f:parse_Continue },
	0x0041: { n:"Pane", f:parse_Pane },
	0x004d: { n:"Pls", f:parse_Pls },
	0x0050: { n:"DCon", f:parse_DCon },
	0x0051: { n:"DConRef", f:parse_DConRef },
	0x0052: { n:"DConName", f:parse_DConName },
	0x0059: { n:"XCT", f:parse_XCT },
	0x005a: { n:"CRN", f:parse_CRN },
	0x005b: { n:"FileSharing", f:parse_FileSharing },
	0x005d: { n:"Obj", f:parse_Obj },
	0x005e: { n:"Uncalced", f:parse_Uncalced },
	0x0060: { n:"Template", f:parse_Template },
	0x0061: { n:"Intl", f:parse_Intl },
	0x007d: { n:"ColInfo", f:parse_ColInfo },
	0x0080: { n:"Guts", f:parse_Guts },
	0x0081: { n:"WsBool", f:parse_WsBool },
	0x0086: { n:"WriteProtect", f:parse_WriteProtect },
	0x0090: { n:"Sort", f:parse_Sort },
	0x0092: { n:"Palette", f:parse_Palette },
	0x0097: { n:"Sync", f:parse_Sync },
	0x0098: { n:"LPr", f:parse_LPr },
	0x0099: { n:"DxGCol", f:parse_DxGCol },
	0x009a: { n:"FnGroupName", f:parse_FnGroupName },
	0x009b: { n:"FilterMode", f:parse_FilterMode },
	0x009d: { n:"AutoFilterInfo", f:parse_AutoFilterInfo },
	0x009e: { n:"AutoFilter", f:parse_AutoFilter },
	0x00a0: { n:"Scl", f:parse_Scl },
	0x00a1: { n:"Setup", f:parse_Setup },
	0x00ae: { n:"ScenMan", f:parse_ScenMan },
	0x00af: { n:"SCENARIO", f:parse_SCENARIO },
	0x00b0: { n:"SxView", f:parse_SxView },
	0x00b1: { n:"Sxvd", f:parse_Sxvd },
	0x00b2: { n:"SXVI", f:parse_SXVI },
	0x00b4: { n:"SxIvd", f:parse_SxIvd },
	0x00b5: { n:"SXLI", f:parse_SXLI },
	0x00b6: { n:"SXPI", f:parse_SXPI },
	0x00b8: { n:"DocRoute", f:parse_DocRoute },
	0x00b9: { n:"RecipName", f:parse_RecipName },
	0x00be: { n:"MulBlank", f:parse_MulBlank },
	0x00c5: { n:"SXDI", f:parse_SXDI },
	0x00c6: { n:"SXDB", f:parse_SXDB },
	0x00c7: { n:"SXFDB", f:parse_SXFDB },
	0x00c8: { n:"SXDBB", f:parse_SXDBB },
	0x00c9: { n:"SXNum", f:parse_SXNum },
	0x00cb: { n:"SxErr", f:parse_SxErr },
	0x00cc: { n:"SXInt", f:parse_SXInt },
	0x00cd: { n:"SXString", f:parse_SXString },
	0x00ce: { n:"SXDtr", f:parse_SXDtr },
	0x00cf: { n:"SxNil", f:parse_SxNil },
	0x00d0: { n:"SXTbl", f:parse_SXTbl },
	0x00d1: { n:"SXTBRGIITM", f:parse_SXTBRGIITM },
	0x00d2: { n:"SxTbpg", f:parse_SxTbpg },
	0x00d3: { n:"ObProj", f:parse_ObProj },
	0x00d5: { n:"SXStreamID", f:parse_SXStreamID },
	0x00d7: { n:"DBCell", f:parse_DBCell },
	0x00d8: { n:"SXRng", f:parse_SXRng },
	0x00d9: { n:"SxIsxoper", f:parse_SxIsxoper },
	0x00da: { n:"BookBool", f:parse_BookBool },
	0x00dc: { n:"DbOrParamQry", f:parse_DbOrParamQry },
	0x00de: { n:"OleObjectSize", f:parse_OleObjectSize },
	0x00e0: { n:"XF", f:parse_XF },
	0x00e3: { n:"SXVS", f:parse_SXVS },
	0x00e5: { n:"MergeCells", f:parse_MergeCells },
	0x00e9: { n:"BkHim", f:parse_BkHim },
	0x00eb: { n:"MsoDrawingGroup", f:parse_MsoDrawingGroup },
	0x00ec: { n:"MsoDrawing", f:parse_MsoDrawing },
	0x00ed: { n:"MsoDrawingSelection", f:parse_MsoDrawingSelection },
	0x00ef: { n:"PhoneticInfo", f:parse_PhoneticInfo },
	0x00f0: { n:"SxRule", f:parse_SxRule },
	0x00f1: { n:"SXEx", f:parse_SXEx },
	0x00f2: { n:"SxFilt", f:parse_SxFilt },
	0x00f4: { n:"SxDXF", f:parse_SxDXF },
	0x00f5: { n:"SxItm", f:parse_SxItm },
	0x00f6: { n:"SxName", f:parse_SxName },
	0x00f7: { n:"SxSelect", f:parse_SxSelect },
	0x00f8: { n:"SXPair", f:parse_SXPair },
	0x00f9: { n:"SxFmla", f:parse_SxFmla },
	0x00fb: { n:"SxFormat", f:parse_SxFormat },
	0x00ff: { n:"ExtSST", f:parse_ExtSST },
	0x0100: { n:"SXVDEx", f:parse_SXVDEx },
	0x0103: { n:"SXFormula", f:parse_SXFormula },
	0x0122: { n:"SXDBEx", f:parse_SXDBEx },
	0x0137: { n:"RRDInsDel", f:parse_RRDInsDel },
	0x0138: { n:"RRDHead", f:parse_RRDHead },
	0x013b: { n:"RRDChgCell", f:parse_RRDChgCell },
	0x013e: { n:"RRDRenSheet", f:parse_RRDRenSheet },
	0x013f: { n:"RRSort", f:parse_RRSort },
	0x0140: { n:"RRDMove", f:parse_RRDMove },
	0x014a: { n:"RRFormat", f:parse_RRFormat },
	0x014b: { n:"RRAutoFmt", f:parse_RRAutoFmt },
	0x014d: { n:"RRInsertSh", f:parse_RRInsertSh },
	0x014e: { n:"RRDMoveBegin", f:parse_RRDMoveBegin },
	0x014f: { n:"RRDMoveEnd", f:parse_RRDMoveEnd },
	0x0150: { n:"RRDInsDelBegin", f:parse_RRDInsDelBegin },
	0x0151: { n:"RRDInsDelEnd", f:parse_RRDInsDelEnd },
	0x0152: { n:"RRDConflict", f:parse_RRDConflict },
	0x0153: { n:"RRDDefName", f:parse_RRDDefName },
	0x0154: { n:"RRDRstEtxp", f:parse_RRDRstEtxp },
	0x015f: { n:"LRng", f:parse_LRng },
	0x0191: { n:"CUsr", f:parse_CUsr },
	0x0192: { n:"CbUsr", f:parse_CbUsr },
	0x0193: { n:"UsrInfo", f:parse_UsrInfo },
	0x0194: { n:"UsrExcl", f:parse_UsrExcl },
	0x0195: { n:"FileLock", f:parse_FileLock },
	0x0196: { n:"RRDInfo", f:parse_RRDInfo },
	0x0197: { n:"BCUsrs", f:parse_BCUsrs },
	0x0198: { n:"UsrChk", f:parse_UsrChk },
	0x01a9: { n:"UserBView", f:parse_UserBView },
	0x01aa: { n:"UserSViewBegin", f:parse_UserSViewBegin },
	0x01ab: { n:"UserSViewEnd", f:parse_UserSViewEnd },
	0x01ac: { n:"RRDUserView", f:parse_RRDUserView },
	0x01ad: { n:"Qsi", f:parse_Qsi },
	0x01ae: { n:"SupBook", f:parse_SupBook },
	0x01b0: { n:"CondFmt", f:parse_CondFmt },
	0x01b1: { n:"CF", f:parse_CF },
	0x01b2: { n:"DVal", f:parse_DVal },
	0x01b5: { n:"DConBin", f:parse_DConBin },
	0x01b6: { n:"TxO", f:parse_TxO },
	0x01b8: { n:"HLink", f:parse_HLink },
	0x01b9: { n:"Lel", f:parse_Lel },
	0x01ba: { n:"CodeName", f:parse_CodeName },
	0x01bb: { n:"SXFDBType", f:parse_SXFDBType },
	0x01bd: { n:"ObNoMacros", f:parse_ObNoMacros },
	0x01be: { n:"Dv", f:parse_Dv },
	0x0201: { n:"Blank", f:parse_Blank },
	0x0203: { n:"Number", f:parse_Number },
	0x0204: { n:"Label", f:parse_Label },
	0x0205: { n:"BoolErr", f:parse_BoolErr },
	0x020b: { n:"Index", f:parse_Index },
	0x0221: { n:"Array", f:parse_Array },
	0x0236: { n:"Table", f:parse_Table },
	0x023e: { n:"Window2", f:parse_Window2 },
	0x027e: { n:"RK", f:parse_RK },
	0x0293: { n:"Style", f:parse_Style },
	0x0418: { n:"BigName", f:parse_BigName },
	0x043c: { n:"ContinueBigName", f:parse_ContinueBigName },
	0x04bc: { n:"ShrFmla", f:parse_ShrFmla },
	0x0800: { n:"HLinkTooltip", f:parse_HLinkTooltip },
	0x0801: { n:"WebPub", f:parse_WebPub },
	0x0802: { n:"QsiSXTag", f:parse_QsiSXTag },
	0x0803: { n:"DBQueryExt", f:parse_DBQueryExt },
	0x0804: { n:"ExtString", f:parse_ExtString },
	0x0805: { n:"TxtQry", f:parse_TxtQry },
	0x0806: { n:"Qsir", f:parse_Qsir },
	0x0807: { n:"Qsif", f:parse_Qsif },
	0x0808: { n:"RRDTQSIF", f:parse_RRDTQSIF },
	0x080a: { n:"OleDbConn", f:parse_OleDbConn },
	0x080b: { n:"WOpt", f:parse_WOpt },
	0x080c: { n:"SXViewEx", f:parse_SXViewEx },
	0x080d: { n:"SXTH", f:parse_SXTH },
	0x080e: { n:"SXPIEx", f:parse_SXPIEx },
	0x080f: { n:"SXVDTEx", f:parse_SXVDTEx },
	0x0810: { n:"SXViewEx9", f:parse_SXViewEx9 },
	0x0812: { n:"ContinueFrt", f:parse_ContinueFrt },
	0x0813: { n:"RealTimeData", f:parse_RealTimeData },
	0x0850: { n:"ChartFrtInfo", f:parse_ChartFrtInfo },
	0x0851: { n:"FrtWrapper", f:parse_FrtWrapper },
	0x0852: { n:"StartBlock", f:parse_StartBlock },
	0x0853: { n:"EndBlock", f:parse_EndBlock },
	0x0854: { n:"StartObject", f:parse_StartObject },
	0x0855: { n:"EndObject", f:parse_EndObject },
	0x0856: { n:"CatLab", f:parse_CatLab },
	0x0857: { n:"YMult", f:parse_YMult },
	0x0858: { n:"SXViewLink", f:parse_SXViewLink },
	0x0859: { n:"PivotChartBits", f:parse_PivotChartBits },
	0x085a: { n:"FrtFontList", f:parse_FrtFontList },
	0x0862: { n:"SheetExt", f:parse_SheetExt },
	0x0863: { n:"BookExt", f:parse_BookExt, r:12},
	0x0864: { n:"SXAddl", f:parse_SXAddl },
	0x0865: { n:"CrErr", f:parse_CrErr },
	0x0866: { n:"HFPicture", f:parse_HFPicture },
	0x0868: { n:"Feat", f:parse_Feat },
	0x086a: { n:"DataLabExt", f:parse_DataLabExt },
	0x086b: { n:"DataLabExtContents", f:parse_DataLabExtContents },
	0x086c: { n:"CellWatch", f:parse_CellWatch },
	0x0871: { n:"FeatHdr11", f:parse_FeatHdr11 },
	0x0872: { n:"Feature11", f:parse_Feature11 },
	0x0874: { n:"DropDownObjIds", f:parse_DropDownObjIds },
	0x0875: { n:"ContinueFrt11", f:parse_ContinueFrt11 },
	0x0876: { n:"DConn", f:parse_DConn },
	0x0877: { n:"List12", f:parse_List12 },
	0x0878: { n:"Feature12", f:parse_Feature12 },
	0x0879: { n:"CondFmt12", f:parse_CondFmt12 },
	0x087a: { n:"CF12", f:parse_CF12 },
	0x087b: { n:"CFEx", f:parse_CFEx },
	0x087c: { n:"XFCRC", f:parse_XFCRC },
	0x087d: { n:"XFExt", f:parse_XFExt },
	0x087e: { n:"AutoFilter12", f:parse_AutoFilter12 },
	0x087f: { n:"ContinueFrt12", f:parse_ContinueFrt12 },
	0x0884: { n:"MDTInfo", f:parse_MDTInfo },
	0x0885: { n:"MDXStr", f:parse_MDXStr },
	0x0886: { n:"MDXTuple", f:parse_MDXTuple },
	0x0887: { n:"MDXSet", f:parse_MDXSet },
	0x0888: { n:"MDXProp", f:parse_MDXProp },
	0x0889: { n:"MDXKPI", f:parse_MDXKPI },
	0x088a: { n:"MDB", f:parse_MDB },
	0x088b: { n:"PLV", f:parse_PLV },
	0x088c: { n:"Compat12", f:parse_Compat12, r:12 },
	0x088d: { n:"DXF", f:parse_DXF },
	0x088e: { n:"TableStyles", f:parse_TableStyles, r:12 },
	0x088f: { n:"TableStyle", f:parse_TableStyle },
	0x0890: { n:"TableStyleElement", f:parse_TableStyleElement },
	0x0892: { n:"StyleExt", f:parse_StyleExt },
	0x0893: { n:"NamePublish", f:parse_NamePublish },
	0x0894: { n:"NameCmt", f:parse_NameCmt },
	0x0895: { n:"SortData", f:parse_SortData },
	0x0896: { n:"Theme", f:parse_Theme },
	0x0897: { n:"GUIDTypeLib", f:parse_GUIDTypeLib },
	0x0898: { n:"FnGrp12", f:parse_FnGrp12 },
	0x0899: { n:"NameFnGrp12", f:parse_NameFnGrp12 },
	0x089a: { n:"MTRSettings", f:parse_MTRSettings },
	0x089c: { n:"HeaderFooter", f:parse_HeaderFooter },
	0x089d: { n:"CrtLayout12", f:parse_CrtLayout12 },
	0x089e: { n:"CrtMlFrt", f:parse_CrtMlFrt },
	0x089f: { n:"CrtMlFrtContinue", f:parse_CrtMlFrtContinue },
	0x08a4: { n:"ShapePropsStream", f:parse_ShapePropsStream },
	0x08a5: { n:"TextPropsStream", f:parse_TextPropsStream },
	0x08a6: { n:"RichTextStream", f:parse_RichTextStream },
	0x08a7: { n:"CrtLayout12A", f:parse_CrtLayout12A },
	0x1001: { n:"Units", f:parse_Units },
	0x1002: { n:"Chart", f:parse_Chart },
	0x1003: { n:"Series", f:parse_Series },
	0x1006: { n:"DataFormat", f:parse_DataFormat },
	0x1007: { n:"LineFormat", f:parse_LineFormat },
	0x1009: { n:"MarkerFormat", f:parse_MarkerFormat },
	0x100a: { n:"AreaFormat", f:parse_AreaFormat },
	0x100b: { n:"PieFormat", f:parse_PieFormat },
	0x100c: { n:"AttachedLabel", f:parse_AttachedLabel },
	0x100d: { n:"SeriesText", f:parse_SeriesText },
	0x1014: { n:"ChartFormat", f:parse_ChartFormat },
	0x1015: { n:"Legend", f:parse_Legend },
	0x1016: { n:"SeriesList", f:parse_SeriesList },
	0x1017: { n:"Bar", f:parse_Bar },
	0x1018: { n:"Line", f:parse_Line },
	0x1019: { n:"Pie", f:parse_Pie },
	0x101a: { n:"Area", f:parse_Area },
	0x101b: { n:"Scatter", f:parse_Scatter },
	0x101c: { n:"CrtLine", f:parse_CrtLine },
	0x101d: { n:"Axis", f:parse_Axis },
	0x101e: { n:"Tick", f:parse_Tick },
	0x101f: { n:"ValueRange", f:parse_ValueRange },
	0x1020: { n:"CatSerRange", f:parse_CatSerRange },
	0x1021: { n:"AxisLine", f:parse_AxisLine },
	0x1022: { n:"CrtLink", f:parse_CrtLink },
	0x1024: { n:"DefaultText", f:parse_DefaultText },
	0x1025: { n:"Text", f:parse_Text },
	0x1027: { n:"ObjectLink", f:parse_ObjectLink },
	0x1032: { n:"Frame", f:parse_Frame },
	0x1033: { n:"Begin", f:parse_Begin },
	0x1034: { n:"End", f:parse_End },
	0x1035: { n:"PlotArea", f:parse_PlotArea },
	0x103a: { n:"Chart3d", f:parse_Chart3d },
	0x103c: { n:"PicF", f:parse_PicF },
	0x103d: { n:"DropBar", f:parse_DropBar },
	0x103e: { n:"Radar", f:parse_Radar },
	0x103f: { n:"Surf", f:parse_Surf },
	0x1040: { n:"RadarArea", f:parse_RadarArea },
	0x1041: { n:"AxisParent", f:parse_AxisParent },
	0x1043: { n:"LegendException", f:parse_LegendException },
	0x1044: { n:"ShtProps", f:parse_ShtProps },
	0x1045: { n:"SerToCrt", f:parse_SerToCrt },
	0x1046: { n:"AxesUsed", f:parse_AxesUsed },
	0x1048: { n:"SBaseRef", f:parse_SBaseRef },
	0x104a: { n:"SerParent", f:parse_SerParent },
	0x104b: { n:"SerAuxTrend", f:parse_SerAuxTrend },
	0x104e: { n:"IFmtRecord", f:parse_IFmtRecord },
	0x104f: { n:"Pos", f:parse_Pos },
	0x1050: { n:"AlRuns", f:parse_AlRuns },
	0x1051: { n:"BRAI", f:parse_BRAI },
	0x105b: { n:"SerAuxErrBar", f:parse_SerAuxErrBar },
	0x105c: { n:"ClrtClient", f:parse_ClrtClient },
	0x105d: { n:"SerFmt", f:parse_SerFmt },
	0x105f: { n:"Chart3DBarShape", f:parse_Chart3DBarShape },
	0x1060: { n:"Fbi", f:parse_Fbi },
	0x1061: { n:"BopPop", f:parse_BopPop },
	0x1062: { n:"AxcExt", f:parse_AxcExt },
	0x1063: { n:"Dat", f:parse_Dat },
	0x1064: { n:"PlotGrowth", f:parse_PlotGrowth },
	0x1065: { n:"SIIndex", f:parse_SIIndex },
	0x1066: { n:"GelFrame", f:parse_GelFrame },
	0x1067: { n:"BopPopCustom", f:parse_BopPopCustom },
	0x1068: { n:"Fbi2", f:parse_Fbi2 },
	0x0000: {}
};
