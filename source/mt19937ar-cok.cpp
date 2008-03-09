/* 
   C++ functions for MT19937, with initialization improved 2002/2/10.
   Coded by Takuji Nishimura and Makoto Matsumoto.
   This is a faster version by taking Shawn Cokus's optimization,
   Matthe Bellew's simplification, Isaku Wada's real version.

   Before using, initialize the state by using init_genrand(seed) 
   or init_by_array(init_key, key_length).

   Copyright (C) 1997 - 2002, Makoto Matsumoto and Takuji Nishimura,
   All rights reserved.                          

   Redistribution and use in source and binary forms, with or without
   modification, are permitted provided that the following conditions
   are met:

     1. Redistributions of source code must retain the above copyright
        notice, this list of conditions and the following disclaimer.

     2. Redistributions in binary form must reproduce the above copyright
        notice, this list of conditions and the following disclaimer in the
        documentation and/or other materials provided with the distribution.

     3. The names of its contributors may not be used to endorse or promote 
        products derived from this software without specific prior written 
        permission.

   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
   A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER OR
   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


   Any feedback is very welcome.
   http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
   email: m-mat@math.sci.hiroshima-u.ac.jp
*/

#include "StdAfx.h"								// Pre-compiled headers
//#include <time.h> // Don't rely on any of these functions due to their code size.
#include "mt19937ar-cok.h"

// Period parameters 
#define N 624
#define M 397
#define MATRIX_A 0x9908b0dfUL   // constant vector a 
#define UMASK 0x80000000UL // most significant w-r bits
#define LMASK 0x7fffffffUL // least significant r bits 
#define MIXBITS(u,v) ( ((u) & UMASK) | ((v) & LMASK) )
#define TWIST(u,v) ((MIXBITS(u,v) >> 1) ^ ((v)&1UL ? MATRIX_A : 0UL))

static unsigned long state[N]; // the array for the state vector 
static int left = 1;
static int initf = 0;
static unsigned long *next;

// initializes state[N] with a seed 
void init_genrand(unsigned long s)
{
    int j;

    state[0]= s & 0xffffffffUL;
    for (j=1; j<N; j++) {
        /* See Knuth TAOCP Vol2. 3rd Ed. P.106 for multiplier.
         * In the previous versions, MSBs of the seed affect  
         * only MSBs of the array state[].                    
         * 2002/01/09 modified by Makoto Matsumoto             */
        state[j] = (1812433253UL * (state[j-1] ^ (state[j-1] >> 30)) + j);
        state[j] &= 0xffffffffUL;  // for >32 bit machines 
    }
    left = 1;
	initf = 1;
}


// AutoHotkey: Comment out unused functions because sometimes the compiler puts them in the EXE.  This happened
// with PCRE but not with the functions in this particular file; but even so, it seems best to do so in case
// other compilers behave differently.

///* initialize by an array with array-length
// * init_key is the array for initializing keys 
// * key_length is its length */
//void init_by_array(unsigned long init_key[], long key_length)
//{
//    int i, j, k;
//
//    init_genrand(19650218UL);
//    i=1;
//	j=0;
//    k = (N>key_length ? N : key_length);  //return lesser of N and key_length
//    for (; k; k--) {
//        state[i] = (state[i] ^ ((state[i-1] ^ (state[i-1] >> 30)) * 1664525UL)) + init_key[j] + j; // non linear 
//        state[i] &= 0xffffffffUL; // for WORDSIZE > 32 machines
//        i++;
//		j++;
//        if (i>=N) { 
//			state[0] = state[N-1];
//			i=1;
//		}
//        if (j>=key_length) 
//			j = 0;
//    }
//    for (k=N-1; k; k--) {
//        state[i] = (state[i] ^ ((state[i-1] ^ (state[i-1] >> 30)) * 1566083941UL)) - i; // non linear 
//        state[i] &= 0xffffffffUL; // for WORDSIZE > 32 machines 
//        i++;
//        if (i>=N) { 
//			state[0] = state[N-1];
//			i=1;
//		}
//    }
//
//    state[0] = 0x80000000UL; // MSB is 1; assuring non-zero initial array 
//    left = 1;
//	initf = 1;
//}

static void next_state(void)
{
    unsigned long *p=state;
    int j;

    // if init_genrand() has not been called, a default initial seed is used     
    if (initf==0) 
		//init_genrand((unsigned long)time(NULL));  // get timer value
		init_genrand(GetTickCount()); // Avoiding the c-lib time functions reduces code size for AHK.

    left = N;
    next = state;
    
    for (j=N-M+1; --j; p++) 
        *p = p[M] ^ TWIST(p[0], p[1]);

    for (j=M; --j; p++) 
        *p = p[M-N] ^ TWIST(p[0], p[1]);

    *p = p[M-N] ^ TWIST(p[0], state[0]);
}

// generates a random number on [0,0xffffffff]-interval
unsigned long genrand_int32(void)
{
    unsigned long y;

    if (--left == 0) 
		next_state();
    y = *next++;

    // Tempering 
    y ^= (y >> 11);
    y ^= (y << 7) & 0x9d2c5680UL;
    y ^= (y << 15) & 0xefc60000UL;
    y ^= (y >> 18);

    return y;
}

// generates a random number on [0,0x7fffffff]-interval
long genrand_int31(void)
{
    unsigned long y;

    if (--left == 0) 
		next_state();
    y = *next++;

    // Tempering 
    y ^= (y >> 11);
    y ^= (y << 7) & 0x9d2c5680UL;
    y ^= (y << 15) & 0xefc60000UL;
    y ^= (y >> 18);

    return (long)(y>>1);
}

// generates a random number on [0,1]-real-interval 
double genrand_real1(void)
{
    unsigned long y;

    if (--left == 0) 
		next_state();
    y = *next++;

    // Tempering 
    y ^= (y >> 11);
    y ^= (y << 7) & 0x9d2c5680UL;
    y ^= (y << 15) & 0xefc60000UL;
    y ^= (y >> 18);

    return (double)y * (1.0/4294967295.0);
    // divided by 2^32-1 
}


// AutoHotkey: Comment out unused functions (see similar comment above for details).

//double genrand_real2(void)
//{
//    unsigned long y;
//
//    if (--left == 0) 
//		next_state();
//    y = *next++;
//
//    // Tempering 
//    y ^= (y >> 11);
//    y ^= (y << 7) & 0x9d2c5680UL;
//    y ^= (y << 15) & 0xefc60000UL;
//    y ^= (y >> 18);
//
//    return (double)y * (1.0/4294967296.0);
//    // divided by 2^32
//}
//
//// generates a random number on (0,1)-real-interval 
//double genrand_real3(void)
//{
//    unsigned long y;
//
//    if (--left == 0) 
//		next_state();
//    y = *next++;
//
//    // Tempering
//    y ^= (y >> 11);
//    y ^= (y << 7) & 0x9d2c5680UL;
//    y ^= (y << 15) & 0xefc60000UL;
//    y ^= (y >> 18);
//
//    return ((double)y + 0.5) * (1.0/4294967296.0);
//    // divided by 2^32
//}
//
//// generates a random number on [0,1) with 53-bit resolution
//double genrand_res53(void) 
//{ 
//    unsigned long a=genrand_int32()>>5, b=genrand_int32()>>6;
//
//    return(a*67108864.0+b)*(1.0/9007199254740992.0);
//} 
//// These real versions are due to Isaku Wada, 2002/01/09 added */
