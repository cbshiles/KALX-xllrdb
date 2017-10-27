// xllrdb.cpp - Use named ranges as a database.
/*
A rdb is a two column range of key value pairs. Values may be rdb's also.
If the first entry begins with equal "=" then it should be followed with a function name
and the keys should correpsond to the function arguments.
*/
#include "xllrdb.h"
#include "xll/utility/uuid.h"

using namespace xll;
using namespace rdb;

typedef traits<XLOPERX>::xstring xstring;

// prepend a "_"
inline OPERX underscore(const OPERX& o)
{
	return ExcelX(xlfConcatenate, OPERX(_T("_")), o);
}
// counted string to name safe to use with DEFINE.NAME
inline OPERX safe(xcstr name)
{
	OPERX xSafe(name + 1, name[0]);

	if (!_istalpha(name[1]) && name[1] != _T('_'))
		xSafe = underscore(xSafe);

	for (xword i = 0; i < xSafe.val.str[0]; ++i)
		// preserve periods
		if (!_istalnum(xSafe.val.str[i + 1]) && xSafe.val.str[i + 1] != _T('.'))
			xSafe.val.str[i + 1] = _T('_');

	return xSafe;
}
inline OPERX safe(const OPERX& xName)
{
	ensure (xName.xltype == xltypeStr);

	return safe(xName.val.str);
}
inline OPERX safe(HANDLEX h)
{
	return safe(OPERX(ExcelX(xlfText, OPERX(h), OPERX(_T("General")))));
}
inline OPERX uuid(void)
{
	xstring id = Uuid::String(Uuid::Uuid());
	for (size_t i = 0; i < id.size(); ++i)
		if (id[i] == _T('-'))
			id[i] = _T('_');
	// create hidden name based on UUID
	return underscore(OPERX(id.c_str()));
}
/*
static AddInX xai_rdb_safename(
	FunctionX(XLL_LPOPERX, _T("?xll_rdb_safename"), _T("RDB.SAFENAME"))
	.Arg(XLL_PSTRINGX, _T("String"), _T("is a string to be used to name a range. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Defined names must start with a letter or underscore "))
	.Documentation(
		_T("Defined names must start with a letter or underscore and contain only alphanumeric ")
		_T("letters, underscore \"_\", or period \".\"." )
		_T("All verboten characters are converted to underscore. ")
	)
);
LPOPERX WINAPI
xll_rdb_safename(xcstr s)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = safe(s);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrX(xlerrValue);
	}

	return &o;
}
*/
// recursively undefine a range
static void
undefine(const OPERX& xName)
{
	OPERX xVal = eval(bang(xName));

	if (!xVal)
		return; // not defined

	if (is_rdb(xVal)) {
		for (xword i = 1; i < xVal.rows(); ++i) {
			const OPERX& x = xVal(i,1);
			if (is_id(x))
				undefine(x);
		}
	}

	ExcelX(xlcDeleteName, xName);
}

static OPERX
rdb_define(OPERX& xKV, OPERX xPre)
{
	ensure (is_rdb(xKV));

	if (xPre) {
		xPre = ExcelX(xlfConcatenate, xPre, OPERX(_T(".")), xKV[0]);
	}
	else {
		undefine(xKV[0]);
		xPre = safe(xKV[0]);
	}

	//	extract fomula name
	OPERX oFor = ExcelX(xlfGetCell, OPERX(6), ExcelX(xlfOffset, ExcelX(xlfActiveCell), OPERX(0), OPERX(1))); // formula
	if (oFor.xltype == xltypeStr) {
		OPERX oFind = ExcelX(xlfFind, OPERX(_T("(")), oFor);
		if (oFind.xltype == xltypeNum) {
			oFor = ExcelX(xlfRight, ExcelX(xlfLeft, oFor, OPERX(oFind - 1)), OPERX(oFind - 2));
			xKV[1] = oFor;
		}
	}

	for (xword i = 1; i < xKV.rows(); ++i) {
		OPERX xKey = xKV(i,0);
		OPERX xVal = xKV(i,1);
		if (xVal.xltype == xltypeNil)
			xKV(i,1) = 0; // fix up blank cell to zero, just like Excel.
		if (xVal.xltype == xltypeNum) {
			xll::handle<OPERX> h(xVal.val.num);
			if (h) {
				OPERX xId = uuid();
				xKV(i, 1) = xId;
				if (is_rdb(*h))
					ExcelX(xlcDefineName, rdb_define(*h, xPre), xId);
#ifdef _DEBUG
				ExcelX(xlcDefineName, xId, *h);
//				ExcelX(xlcDefineName, xId, *h, OPERX(3), MissingX(), OPERX(true)); // hide
#else
				ExcelX(xlcDefineName, xId, *h, OPERX(3), MissingX(), OPERX(true)); // hide
#endif
			}
		}
	}

	ensure (ExcelX(xlcDefineName, xPre, xKV));

	return xPre;
}

static AddInX xai_rdb_define(
	MacroX(_T("_xll_rdb_define@0"), _T("RDB.DEFINE"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Defines a named range that can be used as a simple database. Shortcut Ctrl-Shift-D."))
	.Documentation(
		_T("The range should be two columns of key-value pairs. ")
	)
);
extern "C" int __declspec(dllexport) WINAPI
xll_rdb_define(void)
{
	OPERX xCalc;

	try {
		static OPERX xPre(_T(""));

		// get current calculation mode
		xCalc = ExcelX(xlfGetDocument, OPERX(14));
		ExcelX(xlcOptionsCalculation, OPERX(3)); // manual

		OPERX xSel = ExcelX(xlCoerce, ExcelX(xlfSelection));
		rdb_define(xSel, xPre);

		ExcelX(xlcOptionsCalculation, xCalc);
	}
	catch (const std::exception& ex) {
		ExcelX(xlcOptionsCalculation, xCalc);
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

static AddInX xai_rdb(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_rdb"), _T("RDB"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return database string to be used in RDB.* functions. "))
	.Documentation(_T("Uses Excel links to pull in external databases. "))
);
LPOPERX WINAPI
xll_rdb(void)
{
#pragma XLLEXPORT
	static OPERX o;

//	ensure (ExcelX(xlcOpenLinks, ExcelX(xlfLinks), OPERX(true), OPERX(1))); // open links read-only
	
	o = ExcelX(xlfGetDocument, OPERX(76)); // active workbook as [Book1]Sheet1

	return &o;
}

static AddInX xai_rdb_key(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_rdb_key"), _T("RDB.KEY"))
	.Arg(XLL_PSTRINGX, _T("Key"), _T("is the string key to look up in the first column of rdb. "))
	.Arg(XLL_LPOPERX, _T("RangeDB"), _T("is a rdb created by RDB.DEFINE."))
	.Arg(XLL_LPOPERX, _T("_Dbc"), _T("is an optional handle to an external range database. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the value in the second column of rdb corresponding the the Key in the first column rdb."))
	.Documentation(
		_T("If the value is a named range then it is replaced by the value of the named range. ")
	)
);
LPOPERX WINAPI
xll_rdb_key(xcstr key, LPOPERX pr, LPOPERX pdbc)
{
#pragma XLLEXPORT
	static OPERX v;

	try {
		OPERX r;

		if (pr->xltype == xltypeStr) {
			r = eval(*pr, *pdbc);

			if (r.xltype == xltypeMulti)
				pr = &r;
		}
		else if (pr->xltype == xltypeNum) {
			xll::handle<OPERX> h(pr->val.num);

			if (h)
				pr = xll::h2p<OPERX>(pr->val.num);
		}
		
		if (pr->xltype != xltypeMulti) {
			return 0;
		}

		v = ExcelX(xlfVlookup, OPERX(key + 1, key[0]), *pr, OPERX(2), OPERX(false));

		if (v.xltype == xltypeErr || v.xltype == xltypeNil) {
			v.xltype = xltypeMissing; // so functions taking OPER args work correctly
		}
		else if (is_id(v)) {
			v = eval(v, *pdbc);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &v;
}

static AddInX xai_rdb_value(
	FunctionX(XLL_LPOPERX, _T("?xll_rdb_value"), _T("RDB.VALUE"))
	.Arg(XLL_LPOPERX, _T("Value"), _T("is the value to be looked up."))
	.Arg(XLL_LPOPERX, _T("RangeDB"), _T("is a rdb created by RDB.DEFINE."))
	.Arg(XLL_LPOPERX, _T("_Dbc"), _T("is an optional handle to an external range database. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the key in the first column of rdb corresponding the the Value in the second column rdb."))
	.Documentation(
		_T("If the value is a named range then it is replaced by the value of the named range. ")
	)
);
LPOPERX WINAPI
xll_rdb_value(LPOPERX pv, LPOPERX pr, LPOPERX pdbc)
{
#pragma XLLEXPORT
	static OPERX k;

	try {
		OPERX r;

		if (pr->xltype == xltypeStr) {
			r = eval(*pr, *pdbc);

			if (r.xltype == xltypeMulti)
				pr = &r;
		}
		else if (pr->xltype == xltypeNum) {
			xll::handle<OPERX> h(pr->val.num);

			if (h)
				pr = xll::h2p<OPERX>(pr->val.num);
		}
		
		if (pr->xltype != xltypeMulti) {
			return 0;
		}

		k.xltype = xltypeNil;
		for (xword i = 0; i < rows<XLOPERX>(*pr); ++i) {
			if (index<XLOPERX>(*pr, i, 1) == *pv) {
				k = index<XLOPERX>(*pr, i, 0);
				break;
			}
		}

		if (k.xltype == xltypeErr || k.xltype == xltypeNil) {
			k.xltype = xltypeMissing; // so functions taking OPER args work correctly
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &k;
}

static AddInX xai_rdb_eval(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_rdb_eval"), _T("RDB.EVAL"))
	.Arg(XLL_PSTRINGX, _T("Name"), _T("is the string name of a named range. "))
	.Arg(XLL_LPOPERX, _T("Dbc"), _T("is an optional database connection. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Evaluates the Name and returns the correponding value. "))
	.Documentation()
);
LPOPERX WINAPI
xll_rdb_eval(xcstr name, LPOPERX pdbc)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = eval(OPERX(name + 1, name[0]), *pdbc);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

static AddInX xai_rdb_call(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_rdb_call"), _T("RDB.CALL"))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is a two column array of key-value pairs. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Call a function on data."))
	.Documentation(_T(""))
);
LPOPERX WINAPI
xll_rdb_call(LPOPERX pr)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = call(*pr);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}
/*
int
rdb_define_close(void)
{
	try {
		if (Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Define Range")))
			Excel<XLOPER>(xlfDeleteCommand, OPER(7), OPER(4), OPER("Define Range"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		
		return 0;
	}

	return 1;
}
static Auto<Close> xac_rdb_define(rdb_define_close);

int
rdb_define_open(void)
{
 	try {
		// Try to add just above first menu separator.
		OPER oPos;
		oPos = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("-"));
		oPos = 5;

		OPER oAdj = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Define Range"));
		if (oAdj == Err(xlerrNA)) {
			OPER oAdj(1, 5);
			oAdj(0, 0) = "Define Range";
			oAdj(0, 1) = "RDB.DEFINE";
			oAdj(0, 3) = "Define Range selected in spreadsheet.";
			Excel<XLOPER>(xlfAddCommand, OPER(7), OPER(4), oAdj, oPos);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Open> xao_rdb_define(rdb_define_open);
*/
// Ctrl-Shift-C
static On<Key> xok_paste_create(_T("^+D"), _T("RDB.DEFINE"));
