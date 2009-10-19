/* ==== ArtNouveaU Quick Math ~ qmath.h ==================================

   WATCOM C/C++ Standard Libraries replacement for tiny code.
   Copyright (C) 1997 by the gang at ArtNouveaU, All Rights Reserved.

   Microsoft Visual C/C++ Standard Libraries replacement for tiny code.
   Copyright (C) 2002 by Gabriele Scibilia, All Rights Reserved.

        author           : <G>SZ ~ ArtNouveaU
        file creation    : 30 March 1997
        file description : fast floating point math intrinsics

        revision history :
            (30/03/97) the beginning
            (24/10/97) renamed WATCOM C/C++ tinymath.h
            (26/10/97) added sincos intrinsic
            (30/10/97) added acos and asin intrinsics
            (31/10/97) added ceil and floor ones
            (01/11/97) added hypot, cosh, sinh, acosh, asinh atanh ones
            (02/11/97) added tanh intrinsic
            (09/01/98) added M_ constants
            (27/08/02) renamed MS Visual C/C++ qmath.h
			(11/01/04) modified qmathTan() - J.Bennett <jon@hiddensoft.com>

        references       :
            ArtNouveaU Tiny Library,
            tinymath WATCOM C/C++ intrinsics for math functions,
            Gabriele Scibilia, 24 October 1997.

            Pascal/Cubic math intrinsics, 1997.

            WATCOM C/C++ v11.0 Math Library, 1997.

            RTLIB2 small math run time library replacement,
            "Art" 64k demo by Farbrausch and Scoopex, 2001.

        notes            : pow returns NaN with non-positive arguments

                           For those who hates "naked" convention
                           try modifing each function as the following
                           ie.
                            i- changing float loads
                           ii- removing inlined "ret" statement

                           double _QMATH_LINK qmathSqrt(double __x)
                           {
	                           __asm fld        qword ptr __x
                               __asm fsqrt
                           }

                           When single precision floating point support
                           is required...

                           _QMATH_NAKED _QMATH_INLINE
                           float _QMATH_LINK qmathSinFloat(float __x)
                           {
                               __asm fld       dword ptr [esp + 4]
                               __asm fsin
                               __asm ret       4
                           }

  ======================================================================== */

#ifndef _QUICKMATH_H_INCLUDED
#define _QUICKMATH_H_INCLUDED

#ifdef __cplusplus
extern "C" {
#endif

// Commented out in AutoHotkey v1.0.40.02 to help reduce code size:
//static char quickmath_id[] = "$Id: qmath.h,v 1.1 2004/01/15 19:50:35 jonbennett Exp $";

// for silly C compilers allowing inlining
// try using "static" definition otherwise
#define _QMATH_INLINE __inline

// functions are emitted without prolog or epilog code
#define _QMATH_NAKED  __declspec(naked)

// argument-passing right to left
// called function pops its own arguments from the stack
#define _QMATH_LINK   __stdcall

#define FIST_MAGIC (((65536.0 * 65536.0 * 16) + 32768.0) * 65536.0)
#define FIST_M0131 (65536.0 * 24.0 * 2.0)
#define FIST_M0230 (65536.0 * 24.0 * 4.0)
#define FIST_M0329 (65536.0 * 24.0 * 8.0)
#define FIST_M0428 (65536.0 * 24.0 * 16.0)
#define FIST_M0527 (65536.0 * 24.0 * 32.0)
#define FIST_M0626 (65536.0 * 24.0 * 64.0)
#define FIST_M0725 (65536.0 * 24.0 * 128.0)
#define FIST_M0824 (65536.0 * 24.0 * 256.0)
#define FIST_M0923 (65536.0 * 24.0 * 512.0)
#define FIST_M1022 (65536.0 * 24.0 * 1024.0)
#define FIST_M1121 (65536.0 * 24.0 * 2048.0)
#define FIST_M1220 (65536.0 * 24.0 * 4096.0)
#define FIST_M1319 (65536.0 * 24.0 * 8192.0)
#define FIST_M1418 (65536.0 * 24.0 * 16384.0)
#define FIST_M1517 (65536.0 * 24.0 * 32768.0)
#define FIST_M1616 (65536.0 * 65536.0 * 24.0 * 1.0)
#define FIST_M1715 (65536.0 * 65536.0 * 24.0 * 2.0)
#define FIST_M1814 (65536.0 * 65536.0 * 24.0 * 4.0)
#define FIST_M1913 (65536.0 * 65536.0 * 24.0 * 8.0)
#define FIST_M2012 (65536.0 * 65536.0 * 24.0 * 16.0)
#define FIST_M2111 (65536.0 * 65536.0 * 24.0 * 32.0)
#define FIST_M2210 (65536.0 * 65536.0 * 24.0 * 64.0)
#define FIST_M2309 (65536.0 * 65536.0 * 24.0 * 128.0)
#define FIST_M2408 (65536.0 * 65536.0 * 24.0 * 256.0)
#define FIST_M2507 (65536.0 * 65536.0 * 24.0 * 512.0)
#define FIST_M2606 (65536.0 * 65536.0 * 24.0 * 1024.0)
#define FIST_M2705 (65536.0 * 65536.0 * 24.0 * 2048.0)
#define FIST_M2804 (65536.0 * 65536.0 * 24.0 * 4096.0)
#define FIST_M2903 (65536.0 * 65536.0 * 24.0 * 8192.0)
#define FIST_M3002 (65536.0 * 65536.0 * 24.0 * 16384.0)
#define FIST_M3101 (65536.0 * 65536.0 * 24.0 * 32768.0)
#define FIST_SHORT (14680064.0f)


/*
The following ASM code will convert quickly from a floating point value
to a 16.16 fixed point number.

This code comes complements of Chris Babcock (a.k.a. "VOR")


i_bignum  dq        04238000000000000h

FADD   [i_bignum]
...
FSTP  [qword ptr i_temp]
...
MOV  eax,[dword ptr i_temp]   ; eax = 16.16
*/


_QMATH_INLINE long qmathFtstNeg(float inval)
{
	return ((*(long *) &inval) > 0x80000000);
}


_QMATH_INLINE long qmathFtstPos(float inval)
{
	return ((*(long *) &inval) < 0x80000000);
}


_QMATH_INLINE long qmathFtstZero(float inval)
{
	return ((*(long *) &inval) + (*(long *) &inval));
}


_QMATH_INLINE long qmathFcompGreatThan(float invala, float invalb)
{
	return ((*(long *) &invala) > (*(long *) &invalb));
}


_QMATH_INLINE long qmathFcompLessThan(float invala, float invalb)
{
	return ((*(long *) &invala) < (*(long *) &invalb));
}


_QMATH_INLINE short qmathFistShort(float inval)
{
	float dtemp = FIST_SHORT + inval;

	return ((*(short *) &dtemp) & 0x1FFFFF);
}


_QMATH_INLINE long qmathFistLong(float inval)
{
	double dtemp = FIST_MAGIC + inval;

	return ((*(long *) &dtemp) - 0x80000000);
}


_QMATH_INLINE long qmathFist0131(float inval)
{
	double dtemp = FIST_M0131 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0230(float inval)
{
	double dtemp = FIST_M0230 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0329(float inval)
{
	double dtemp = FIST_M0329 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0428(float inval)
{
	double dtemp = FIST_M0428 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0527(float inval)
{
	double dtemp = FIST_M0527 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0626(float inval)
{
	double dtemp = FIST_M0626 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0725(float inval)
{
	double dtemp = FIST_M0725 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0824(float inval)
{
	double dtemp = FIST_M0824 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist0923(float inval)
{
	double dtemp = FIST_M0923 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1022(float inval)
{
	double dtemp = FIST_M1022 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1121(float inval)
{
	double dtemp = FIST_M1121 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1220(float inval)
{
	double dtemp = FIST_M1220 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1319(float inval)
{
	double dtemp = FIST_M1319 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1418(float inval)
{
	double dtemp = FIST_M1418 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1517(float inval)
{
	double dtemp = FIST_M1517 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1616(float inval)
{
	double dtemp = FIST_M1616 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1715(float inval)
{
	double dtemp = FIST_M1715 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1814(float inval)
{
	double dtemp = FIST_M1814 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist1913(float inval)
{
	double dtemp = FIST_M1913 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2012(float inval)
{
	double dtemp = FIST_M2012 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2111(float inval)
{
	double dtemp = FIST_M2111 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2210(float inval)
{
	double dtemp = FIST_M2210 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2309(float inval)
{
	double dtemp = FIST_M2309 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2408(float inval)
{
	double dtemp = FIST_M2408 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2507(float inval)
{
	double dtemp = FIST_M2507 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2606(float inval)
{
	double dtemp = FIST_M2606 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2705(float inval)
{
	double dtemp = FIST_M2705 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2804(float inval)
{
	double dtemp = FIST_M2804 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist2903(float inval)
{
	double dtemp = FIST_M2903 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist3002(float inval)
{
	double dtemp = FIST_M3002 + inval;

	return (*(long *) &dtemp);
}


_QMATH_INLINE long qmathFist3101(float inval)
{
	double dtemp = FIST_M3101 + inval;

	return (*(long *) &dtemp);
}


// M_ constants
#define M_E        (2.71828182845904523536)
#define M_LOG2E    (1.44269504088896340736)
#define M_LOG10E   (0.434294481903251827651)
#define M_LN2      (0.693147180559945309417)
#define M_LN10     (2.30258509299404568402)
#define M_PI       (3.14159265358979323846)
#define M_PI_2     (1.57079632679489661923)
#define M_PI_4     (0.785398163397448309616)
#define M_1_PI     (0.318309886183790671538)
#define M_2_PI     (0.636619772367581343076)
#define M_1_SQRTPI (0.564189583547756286948)
#define M_2_SQRTPI (1.12837916709551257390)
#define M_SQRT2    (1.41421356237309504880)
#define M_SQRT_2   (0.707106781186547524401)


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathSin(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fsin
	__asm ret		8
}


static double einhalb = 0.5;

_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAsin(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld		st
	__asm fabs
	__asm fcom		dword ptr [einhalb]
	__asm fstsw		ax
	__asm sahf
	__asm jbe		asin_kleiner
	__asm fld1
	__asm fsubrp	st(1), st(0)
	__asm fld		st
	__asm fadd		st(0), st(0)
	__asm fxch		st(1)
	__asm fmul		st(0), st(0)
	__asm fsubp		st(1), st(0)
	__asm jmp		asin_exit

asin_kleiner:

	__asm fstp		st(0)
	__asm fld		st(0)
	__asm fmul		st(0), st(0)
	__asm fld1
	__asm fsubrp	st(1), st(0)

asin_exit:

	__asm fsqrt
	__asm fpatan
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathCos(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fcos
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAcos(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld1
	__asm fchs
	__asm fcomp		st(1)
	__asm fstsw		ax
	__asm je		acos_suckt

	__asm fld		st(0)
	__asm fld1
	__asm fsubrp	st(1), st(0)
	__asm fxch		st(1)
	__asm fld1
	__asm faddp		st(1), st(0)
	__asm fdivp		st(1), st(0)
	__asm fsqrt
	__asm fld1
	__asm jmp		acos_exit

acos_suckt:

	__asm fld1
	__asm fldz

acos_exit:

	__asm fpatan
	__asm fadd		st(0), st(0)
	__asm ret		8
}


/* J.Bennett - added fstp in tan function to remove 1.0 it pushes onto the stack */

_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathTan(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fptan
	__asm fstp		st(0)
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAtan(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld1
	__asm fpatan
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAtan2(double __y, double __x)
{
	__asm fld		qword ptr [esp +  4]
	__asm fld		qword ptr [esp + 12]
	__asm fpatan
	__asm ret		16
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathExp(double __x)
{
	__asm fld		qword ptr [esp + 4];
	__asm fldl2e
	__asm fmulp		st(1), st
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st
	__asm fscale
	__asm fxch
	__asm fstp		st
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathExp2(double __x)
{
	__asm fld		qword ptr [esp + 4];
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st
	__asm fscale
	__asm fxch
	__asm fstp		st
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathExp10(double __x)
{
	__asm fld		qword ptr [esp + 4];
	__asm fldl2t
	__asm fmulp		st(1), st
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st
	__asm fscale
	__asm fxch
	__asm fstp		st
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathLog(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fldln2
	__asm fxch
	__asm fyl2x
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathLog2(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld1
	__asm fxch
	__asm fyl2x
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathLog10(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fldlg2
	__asm fxch
	__asm fyl2x
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathFabs(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fabs
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathPow(double __x, double __y)
{
	__asm fld		qword ptr [esp + 12]
	__asm fld		qword ptr [esp +  4]
	__asm ftst
	__asm fstsw		ax
	__asm sahf
	__asm jz		pow_zero

	__asm fyl2x
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st(0)
	__asm fscale

pow_zero:

	__asm fstp		st(1)
	__asm ret		16
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathCeil(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fchs
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm sub		esp, 4
	__asm fst		dword ptr [esp]
	__asm fxch		st(2)
	__asm mov		eax, [esp]
	__asm cmp		eax, 0x80000000
	__asm jbe		ceil_exit
	__asm fsub		st, st(1)

ceil_exit:

	__asm fsub		st, st(2)
	__asm fstp		st(1)
	__asm fstp		st(1)
	__asm fchs
	__asm pop		eax
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathFloor(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm sub		esp, 4
	__asm fst		dword ptr [esp]
	__asm fxch		st(2)
	__asm mov		eax, [esp]
	__asm cmp		eax, 0x80000000
	__asm jbe		floor_exit
	__asm fsub		st, st(1)

floor_exit:

	__asm fsub		st, st(2)
	__asm fstp		st(1)
	__asm fstp		st(1)
	__asm pop		eax
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathFmod(double __x, double __y)
{
	__asm fld		qword ptr [esp + 12]
	__asm fld		qword ptr [esp +  4]
	__asm fprem
	__asm fxch
	__asm fstp		st
	__asm ret		16
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathSqrt(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fsqrt
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathHypot(double __x, double __y)
{
	__asm fld		qword ptr [esp + 12]
	__asm fld		qword ptr [esp +  4]
	__asm fmul		st, st
	__asm fxch
	__asm fmul		st, st
	__asm faddp		st(1), st
	__asm fsqrt
	__asm ret		16
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAcosh(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld		st
	__asm fmul		st, st
	__asm fld1
	__asm fsubp		st(1), st
	__asm fsqrt
	__asm faddp		st(1), st
	__asm fldln2
	__asm fxch
	__asm fyl2x
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAsinh(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld		st
	__asm fmul		st, st
	__asm fld1
	__asm faddp		st(1), st
	__asm fsqrt
	__asm faddp		st(1), st
	__asm fldln2
	__asm fxch
	__asm fyl2x
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathAtanh(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld1
	__asm fsub		st, st(1)
	__asm fld1
	__asm faddp		st(2), st
	__asm fdivrp	st(1), st
	__asm fldln2
	__asm fxch
	__asm fyl2x
	__asm mov		eax, 0xBF000000
	__asm push		eax
	__asm fld		dword ptr [esp]
	__asm fmulp		st(1), st
	__asm pop		eax
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathCosh(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fldl2e
	__asm fmulp		st(1), st
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st
	__asm fscale
	__asm fxch
	__asm fstp		st
	__asm fld1
	__asm fdiv		st, st(1)
	__asm faddp		st(1), st
	__asm mov		eax, 0x3F000000
	__asm push		eax
	__asm fld		dword ptr [esp]
	__asm fmulp		st(1), st
	__asm pop		eax
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathSinh(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fldl2e
	__asm fmulp		st(1), st
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st
	__asm fscale
	__asm fxch
	__asm fstp		st
	__asm fld1
	__asm fdiv		st, st(1)
	__asm fsubp		st(1), st
	__asm mov		eax, 0x3F000000
	__asm push		eax
	__asm fld		dword ptr [esp]
	__asm fmulp		st(1), st
	__asm pop		eax
	__asm ret		8
}


_QMATH_NAKED _QMATH_INLINE
double _QMATH_LINK qmathTanh(double __x)
{
	__asm fld		qword ptr [esp + 4]
	__asm fld		st
	__asm mov		eax, 0x40000000
	__asm push		eax
	__asm fld		dword ptr [esp]
	__asm fmul		st, st(1)
	__asm fldl2e
	__asm fmulp		st(1), st
	__asm fld1
	__asm fld		st(1)
	__asm fprem
	__asm f2xm1
	__asm faddp		st(1), st
	__asm fscale
	__asm fxch
	__asm fstp		st
	__asm fld1
	__asm fsub		st, st(1)
	__asm fchs
	__asm fld1
	__asm faddp		st(2), st
	__asm fdivrp	st(1), st
	__asm pop		eax
	__asm ret		8
}


#ifdef __cplusplus
  };
#endif

#endif // _QUICKMATH_H_INCLUDED
