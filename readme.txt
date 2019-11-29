OpenVBS
    - The Script Engine for the BASIC-2020 -
---

* What?
    This is a script engine for modern BASIC, compatible with VBScript.

* How?
    > oscript sample.obs



* The Language
    * Statement
        * dim, redim, const
        * if, then, else/elseif, end if
        * do while/until, loop, exit do
        * for each, in, next, exit for
        * goto, @
        * function, end function, exit function, return
            * default param (=)
            * closure
        * class, private, public, default, property let/get, function, end class
        * new
        * me, err
        * on error goto catch, on error catch, on error goto 0
        * option explicit
        * ()
        * {} (JSON object)
        * [] (JSON array)
    * Operator
        * =
        * :=
        * +
        * -
        * ^
        * *
        * /
        * mod
        * &
        * ++
        * --
        * <
        * >
        * <=
        * >=
        * <>
        * ==
        * is
        * not
        * and
        * or
        * xor
        * :, LF
        * .
        * ,
        * ?.
        * ??
        * ===
        * !==
    * Literal
        * true, false, empty, null, nothing
        * Integer number[&H,&B] (64bit signed)
        * Real number (64bit)
        * String ("~")
        * Date (#~#)
    * Embedded var and funcs
        * vbBinaryCompare
        * vbTextCompare
        * vbSunday
        * vbMonday
        * vbTuesday
        * vbWednesday
        * vbThursday
        * vbFriday
        * vbSaturday
        * vbUseSystemDayOfWeek
        * vbFirstJan1
        * vbFirstFourDays
        * vbFirstFullWeek
        * vbGeneralDate
        * vbLongDate
        * vbShortDate
        * vbLongTime
        * vbShortTime
        * vbObjectError
        * vbUseDefault
        * vbTrue
        * vbFalse
        * vbEmpty
        * vbNull
        * vbInteger
        * vbLong
        * vbSingle
        * vbDouble
        * vbCurrency
        * vbDate
        * vbString
        * vbObject
        * vbError
        * vbBoolean
        * vbVariant
        * vbDataObject
        * vbDecimal
        * vbByte
        * vbArray
        * vbCr
        * VbCrLf
        * vbFormFeed
        * vbLf
        * vbNewLine
        * vbNullChar
        * vbNullString
        * vbTab
        * vbVerticalTab
        * CreateObject
        * GetObject
        * TypeName
        * VarType
        * IsArray
        * IsObject
        * IsEmpty
        * IsNull
        * IsDate
        * IsNumeric
        * CBool
        * CByte
        * CCur
        * CDate
        * CDbl
        * CInt
        * CLng
        * CSng
        * CStr
        * LBound
        * UBound
        * Array
        * Randomize
        * Rnd
        * Int
        * Fix
        * Round
        * Sgn
        * Abs
        * Sqr
        * Sin
        * Cos
        * Tan
        * Atn
        * Log
        * Exp
        * Len
        * Left
        * Right
        * Mid
        * LCase
        * UCase
        * LTrim
        * RTrim
        * Trim
        * InStr
        * InStrRev
        * StrComp
        * Replace
        * StrReverse
        * String
        * Space
        * ChrW
        * AscW
        * Oct
        * Hex
        * Join
        * Split
        * Filter
        * Now
        * Date
        * Time
        * Year
        * Month
        * Day
        * Weekday
        * Hour
        * Minute
        * Second
        * Timer
        * DateAdd [YET]
        * DateDiff [YET]
        * DatePart [YET]
        * DateSerial [YET]
        * DateValue [YET]
        * TimeSerial [YET]
        * TimeValue [YET]
        * FormatCurrency [YET]
        * FormatDateTime [YET]
        * FormatNumber [YET]
        * FormatPercent [YET]
    * Other things
        * for, to, step
        * select, case, case else, end select
        * with, end with
        * getref
        * ' (comment)
        * while, wend [DEPRECATED]
        * sub, end sub, exit sub [DEPRECATED]
        * set, call, rem [DEPRECATED]
        * on error resume next [DEPRECATED]



* Change Log
    * 20191129
        [add] source from stdin
        [add] operator: '===' and '!=='
        [add] VarCmp for Empty
        [add] Nullish Coalescing '??'
        [add] Optional Chaining  ‘?.'
        [add] VBS functions on 'osk' file.
    * 20191031
        [fix] '=' problem in 'if' statement
        [add] CLSCTX_LOCAL_SERVER for Windows
        [add] JSON was improved(can Nest, array[], can for-each, from-to string)
        [add] 'goto' by line-head number
        [add] op copy ':='



* License
    CC BY: NaturalStyle PREMIUM
    This work is licensed under a Creative Commons Attribution 4.0 International License.
    http://creativecommons.org/licenses/by/4.0/

---
yrm.20191129
