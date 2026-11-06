//! Excel's classes for function arguments and results, set on the registry's functions in one
//! place: which arguments take one value, a reference or an array; which hand back the result;
//! what the result can be. The evaluator lifts over arrays from them (crate::array_lift), the
//! .xls and .xlsb writers write token classes from them, and a formula from before dynamic
//! arrays gets its `@` from them (crate::legacy).

use super::FunctionRegistry;

/// Excel's class for a function argument, from its function table ([MS-XLS] §2.5.198): what the
/// argument takes, and so how a file writes a reference there.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArgClass {
    /// A reference, kept whole (SUM's, COUNTIF's range). Written reference-class.
    R,
    /// One value: in a formula from before dynamic arrays a range there reads one cell.
    /// Written value-class.
    V,
    /// An array (SUMPRODUCT's, MMULT's). Written array-class.
    A,
    /// A reference read as an array (INDEX's first, LOOKUP's vectors). Written reference-class,
    /// as Excel writes `=INDEX(A1:A3,1)`.
    RA,
}

/// What a function's result can be (see `FunctionDef::returns`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum Returns {
    /// One value, or an array only when an argument that takes one value is given one.
    #[default]
    Value,
    /// An array whatever its arguments: FILTER, SEQUENCE, MMULT.
    Array,
    /// A reference it works out: INDEX, OFFSET, INDIRECT, XLOOKUP.
    Reference,
    /// A number per row or column of its reference: ROW, COLUMN.
    PerCell,
    /// Its last argument's result: LET.
    Last,
}

/// The arguments where a function takes one value. Given an array there, Excel applies the
/// function to each element and returns an array of the results: `LEN(A1:A3)`,
/// `IF(A1:A3>1,A1:A3)`, `ISNUMBER(SEARCH(x,A1:A9))`, `COUNTIF(r,{"a","b"})`, `CHOOSE({1,2},a,b)`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ArgPositions {
    /// Every argument takes arrays as they are.
    #[default]
    None,
    All,
    At(&'static [usize]),
    /// This argument and every one after it.
    From(usize),
    /// The criteria of COUNTIFS (1, 3, 5, …) or of SUMIFS and the like (2, 4, 6, …).
    EveryOtherFrom(usize),
    /// The second of each pair from this argument on, and a last one left over: SWITCH's results
    /// (2, 4, …) and its default.
    PairsFrom(usize),
}

impl ArgPositions {
    pub fn includes(self, position: usize) -> bool {
        match self {
            ArgPositions::None => false,
            ArgPositions::All => true,
            ArgPositions::At(positions) => positions.contains(&position),
            ArgPositions::From(first) => position >= first,
            ArgPositions::EveryOtherFrom(first) => position >= first && (position - first) % 2 == 0,
            ArgPositions::PairsFrom(first) => position > first && (position - first) % 2 == 1,
        }
    }

    /// `includes` in a call with `count` arguments, which places a left-over last one.
    pub fn includes_in(self, position: usize, count: usize) -> bool {
        match self {
            ArgPositions::PairsFrom(first)
                if count > first && (count - first) % 2 == 1 && position + 1 == count =>
            {
                true
            }
            positions => positions.includes(position),
        }
    }
}

/// How a value typed into the formula is read, as opposed to one from a cell or a function:
/// `SUM("5",TRUE)` is 6 and `SUM("x")` is #VALUE!, while text and TRUE in cells are left out.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TypedArgs {
    #[default]
    AsGiven,
    /// SUM, AVERAGE, MIN, MAX and the like: typed text and TRUE/FALSE become numbers.
    Numbers,
    /// AND, OR, XOR: typed "TRUE"/"FALSE" become booleans, other typed text is #VALUE!.
    Logicals,
}

const SCALAR_ARGS: &[(&[&str], ArgPositions)] = &[
    (
        &[
            "IF",
            "IFERROR",
            "IFNA",
            "IFS",
            "CHOOSE",
            "NOT",
            "ISNUMBER",
            "ISTEXT",
            "ISNONTEXT",
            "ISBLANK",
            "ISERROR",
            "ISERR",
            "ISNA",
            "ISLOGICAL",
            "ISEVEN",
            "ISODD",
            "LEFT",
            "RIGHT",
            "MID",
            "LEN",
            "LOWER",
            "UPPER",
            "PROPER",
            "TRIM",
            "CLEAN",
            "VALUE",
            "TEXT",
            "SUBSTITUTE",
            "REPLACE",
            "FIND",
            "SEARCH",
            "EXACT",
            "REPT",
            "CHAR",
            "CODE",
            "UNICODE",
            "UNICHAR",
            "N",
            "T",
            "NUMBERVALUE",
            "DATEVALUE",
            "TIMEVALUE",
            "YEAR",
            "MONTH",
            "DAY",
            "HOUR",
            "MINUTE",
            "SECOND",
            "WEEKDAY",
            "WEEKNUM",
            "ISOWEEKNUM",
            "DATE",
            "TIME",
            "EDATE",
            "EOMONTH",
            "DATEDIF",
            "DAYS",
            "ABS",
            "INT",
            "TRUNC",
            "ROUND",
            "ROUNDUP",
            "ROUNDDOWN",
            "MROUND",
            "MOD",
            "SIGN",
            "SQRT",
            "POWER",
            "EXP",
            "LN",
            "LOG",
            "LOG10",
            "CEILING",
            "FLOOR",
            "EVEN",
            "ODD",
            "QUOTIENT",
            "FACT",
            "DOLLAR",
            "FIXED",
            "ADDRESS",
            "ERROR.TYPE",
            "CONCATENATE",
            "PMT",
            "FV",
            "PV",
            "NPER",
            "RATE",
            "IPMT",
            "PPMT",
        ],
        ArgPositions::All,
    ),
    (&["SWITCH"], ArgPositions::At(&[0])),
    (
        &["COUNTIF", "SUMIF", "AVERAGEIF", "SMALL", "LARGE"],
        ArgPositions::At(&[1]),
    ),
    (
        &["MATCH", "XMATCH", "XLOOKUP", "LOOKUP"],
        ArgPositions::At(&[0]),
    ),
    (&["VLOOKUP", "HLOOKUP"], ArgPositions::At(&[0, 2])),
    (&["INDEX"], ArgPositions::At(&[1, 2])),
    (&["COUNTIFS"], ArgPositions::EveryOtherFrom(1)),
    (
        &["SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS"],
        ArgPositions::EveryOtherFrom(2),
    ),
];

use ArgClass as Class;

/// Excel's argument classes besides one value (V), the default: references (R), arrays (A) and
/// references read as arrays (RA), then the class of the arguments past the list. This is the one
/// place they are set; the .xls and .xlsb writers write a reference's token class from them, and
/// in a formula from before dynamic arrays a range at a V argument reads one cell
/// (crate::legacy). The functions that came with dynamic arrays take arrays everywhere but at
/// their lookup values. Every argument that lifts over arrays (SCALAR_ARGS) is a V argument here
/// except N's and T's, which are R: Excel's table has RO for them, a reference whose first cell
/// they read, so `N(A1:A3)` gets no `@`.
///
/// The classes the .xls writer emits were checked against Excel-authored files: `=INDEX(A1:A3,1)`
/// writes PtgArea(R) and PtgFuncVar(V), and `=SUM(INDEX(A1:A3,1))` writes PtgFuncVar(R), which is
/// why INDEX's first argument is RA (written R) rather than A.
const ARG_CLASSES: &[(&[&str], &[ArgClass], ArgClass)] = &[
    (
        &[
            "SUM",
            "AVERAGE",
            "MIN",
            "MAX",
            "COUNT",
            "COUNTA",
            "PRODUCT",
            "STDEV",
            "VAR",
            "AND",
            "OR",
            "XOR",
            "MEDIAN",
            "AVEDEV",
            "DEVSQ",
            "GEOMEAN",
            "HARMEAN",
            "SUMSQ",
            "KURT",
            "SKEW",
            "SKEW.P",
            "STDEVP",
            "VARP",
            "STDEVA",
            "STDEVPA",
            "VARA",
            "VARPA",
            "AVERAGEA",
            "MAXA",
            "MINA",
            "STDEV.S",
            "STDEV.P",
            "VAR.S",
            "VAR.P",
            "MODE.SNGL",
            "MODE.MULT",
            "DSUM",
            "DCOUNT",
            "DCOUNTA",
            "DAVERAGE",
            "DMAX",
            "DMIN",
            "DGET",
            "DPRODUCT",
            "DSTDEV",
            "DSTDEVP",
            "DVAR",
            "DVARP",
            "COUNTIFS",
            "SUMIFS",
            "AVERAGEIFS",
            "MAXIFS",
            "MINIFS",
            "ROW",
            "COLUMN",
            "ROWS",
            "COLUMNS",
            "AREAS",
            "ISREF",
            "T",
            "N",
            "COUNTBLANK",
            "SHEET",
            "SHEETS",
            "FORMULATEXT",
            "ISFORMULA",
        ],
        &[],
        Class::R,
    ),
    // IF's results and CHOOSE's choices take a value or a reference: a range there is written
    // reference-class, and IF(...,CHOOSE(...),...) writes CHOOSE reference-class (verified
    // against Excel).
    (&["IF"], &[Class::V, Class::R, Class::R], Class::V),
    (&["NPV", "SUBTOTAL", "CHOOSE"], &[Class::V], Class::R),
    (
        &[
            "LARGE",
            "SMALL",
            "QUARTILE",
            "QUARTILE.INC",
            "QUARTILE.EXC",
            "PERCENTILE",
            "PERCENTILE.INC",
            "PERCENTILE.EXC",
            "PERCENTRANK",
            "PERCENTRANK.INC",
            "PERCENTRANK.EXC",
            "TRIMMEAN",
            "ZTEST",
            "COUNTIF",
            "OFFSET",
        ],
        &[Class::R],
        Class::V,
    ),
    (
        &["SUMIF", "AVERAGEIF"],
        &[Class::R, Class::V, Class::R],
        Class::V,
    ),
    // VLOOKUP's table is a reference (verified against Excel's .xlsb: a value-class area there is
    // repaired into a user-defined function call).
    (
        &[
            "RANK", "RANK.EQ", "RANK.AVG", "MATCH", "VLOOKUP", "HLOOKUP", "CELL",
        ],
        &[Class::V, Class::R],
        Class::V,
    ),
    (&["XNPV"], &[Class::V, Class::R, Class::R], Class::V),
    (&["XIRR"], &[Class::R, Class::R], Class::V),
    (
        &["GCD", "LCM", "MULTINOMIAL", "IMSUM", "IMPRODUCT"],
        &[],
        Class::R,
    ),
    (&["SERIESSUM"], &[Class::V, Class::V, Class::V], Class::R),
    (
        &["FVSCHEDULE", "GETPIVOTDATA"],
        &[Class::V, Class::R],
        Class::V,
    ),
    (
        &["FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.STAT"],
        &[Class::V, Class::A, Class::A],
        Class::V,
    ),
    (
        &["FORECAST.ETS.SEASONALITY"],
        &[Class::A, Class::A],
        Class::V,
    ),
    (&["XLOOKUP", "XMATCH"], &[Class::V], Class::A),
    (
        &["NETWORKDAYS", "WORKDAY"],
        &[Class::V, Class::V, Class::R],
        Class::V,
    ),
    (
        &["NETWORKDAYS.INTL", "WORKDAY.INTL"],
        &[Class::V, Class::V, Class::V, Class::R],
        Class::V,
    ),
    (
        &[
            "SUMPRODUCT",
            "MMULT",
            "MDETERM",
            "MINVERSE",
            "TRANSPOSE",
            "SUMX2MY2",
            "SUMX2PY2",
            "SUMXMY2",
            "CORREL",
            "COVAR",
            "COVARIANCE.P",
            "COVARIANCE.S",
            "PEARSON",
            "RSQ",
            "SLOPE",
            "INTERCEPT",
            "STEYX",
            "FTEST",
            "F.TEST",
            "CHITEST",
            "CHISQ.TEST",
            "MODE",
        ],
        &[],
        Class::A,
    ),
    (&["FORECAST", "FORECAST.LINEAR"], &[Class::V], Class::A),
    (
        &["TTEST", "T.TEST", "PROB"],
        &[Class::A, Class::A],
        Class::V,
    ),
    (&["INDEX", "IRR", "MIRR"], &[Class::RA], Class::V),
    (&["LOOKUP"], &[Class::V], Class::RA),
    (&["FREQUENCY"], &[], Class::RA),
    (&["LINEST", "LOGEST"], &[Class::RA, Class::RA], Class::V),
    (
        &["TREND", "GROWTH"],
        &[Class::RA, Class::RA, Class::RA],
        Class::V,
    ),
    // The array form of AGGREGATE (LARGE, SMALL, PERCENTILE, QUARTILE) needs no Ctrl+Shift+Enter.
    (&["AGGREGATE"], &[Class::V, Class::V], Class::A),
    (
        &[
            "LET",
            "LAMBDA",
            "MAP",
            "REDUCE",
            "SCAN",
            "BYROW",
            "BYCOL",
            "MAKEARRAY",
            "ISOMITTED",
            "FILTER",
            "SORT",
            "SORTBY",
            "UNIQUE",
            "SEQUENCE",
            "RANDARRAY",
            "TEXTSPLIT",
            "TEXTBEFORE",
            "TEXTAFTER",
            "VSTACK",
            "HSTACK",
            "TOCOL",
            "TOROW",
            "TAKE",
            "DROP",
            "CHOOSECOLS",
            "CHOOSEROWS",
            "EXPAND",
            "WRAPROWS",
            "WRAPCOLS",
            "TEXTJOIN",
            "CONCAT",
            "PERCENTOF",
            "TRIMRANGE",
            "IMAGE",
        ],
        &[],
        Class::A,
    ),
];

/// The arguments handed back as the result (see `FunctionDef::result_args`).
const RESULT_ARGS: &[(&[&str], ArgPositions)] = &[
    (&["IF"], ArgPositions::At(&[1, 2])),
    (&["IFS"], ArgPositions::EveryOtherFrom(1)),
    (&["SWITCH"], ArgPositions::PairsFrom(1)),
    (&["CHOOSE"], ArgPositions::From(1)),
];

/// What the functions besides one-value ones return (see `Returns`).
const RETURNS: &[(&[&str], Returns)] = &[
    (
        &[
            "FILTER",
            "SORT",
            "SORTBY",
            "UNIQUE",
            "SEQUENCE",
            "RANDARRAY",
            "TRANSPOSE",
            "MMULT",
            "MINVERSE",
            "MUNIT",
            "FREQUENCY",
            "LINEST",
            "LOGEST",
            "TREND",
            "GROWTH",
            "TEXTSPLIT",
            "VSTACK",
            "HSTACK",
            "TOCOL",
            "TOROW",
            "TAKE",
            "DROP",
            "CHOOSECOLS",
            "CHOOSEROWS",
            "EXPAND",
            "WRAPROWS",
            "WRAPCOLS",
            "MODE.MULT",
            "MAP",
            "SCAN",
            "BYROW",
            "BYCOL",
            "MAKEARRAY",
            "TRIMRANGE",
        ],
        Returns::Array,
    ),
    (
        &["INDEX", "OFFSET", "INDIRECT", "XLOOKUP"],
        Returns::Reference,
    ),
    (&["ROW", "COLUMN"], Returns::PerCell),
    (&["LET"], Returns::Last),
];

/// The functions that give one result per range of an array of references.
const PER_RANGE: &[&str] = &[
    "SUBTOTAL",
    "SUMIF",
    "SUMIFS",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "MAXIFS",
    "MINIFS",
];

const TYPED_ARGS: &[(&[&str], TypedArgs)] = &[
    (
        &[
            "SUM", "AVERAGE", "MIN", "MAX", "PRODUCT", "SUMSQ", "MEDIAN", "STDEV", "STDEV.S",
            "STDEV.P", "STDEVP", "VAR", "VAR.S", "VAR.P", "VARP",
        ],
        TypedArgs::Numbers,
    ),
    (&["AND", "OR", "XOR"], TypedArgs::Logicals),
];

/// Sets the classes on the registry's functions.
pub(super) fn apply(registry: &mut FunctionRegistry) {
    for (names, rule) in SCALAR_ARGS {
        for name in *names {
            registry.function_mut(name).scalar_args = *rule;
        }
    }
    for (names, rule) in TYPED_ARGS {
        for name in *names {
            registry.function_mut(name).typed_args = *rule;
        }
    }
    for (names, classes, default) in ARG_CLASSES {
        for name in *names {
            let def = registry.function_mut(name);
            def.arg_classes = classes;
            def.default_arg_class = *default;
        }
    }
    for (names, rule) in RESULT_ARGS {
        for name in *names {
            registry.function_mut(name).result_args = *rule;
        }
    }
    for (names, returns) in RETURNS {
        for name in *names {
            registry.function_mut(name).returns = *returns;
        }
    }
    for name in PER_RANGE {
        registry.function_mut(name).per_range = true;
    }
}
