/*************************************************
*      Perl-Compatible Regular Expressions       *
*************************************************/

/* PCRE is a library of functions to support regular expressions whose syntax
and semantics are as close as possible to those of the Perl 5 language.

                       Written by Philip Hazel
           Copyright (c) 1997-2011 University of Cambridge

-----------------------------------------------------------------------------
Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright notice,
      this list of conditions and the following disclaimer.

    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.

    * Neither the name of the University of Cambridge nor the names of its
      contributors may be used to endorse or promote products derived from
      this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
-----------------------------------------------------------------------------
*/


/* This module contains pcre_exec(), the externally visible function that does
pattern matching using an NFA algorithm, trying to mimic Perl as closely as
possible. There are also some static supporting functions. */

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#define NLBLOCK md             /* Block containing newline information */
#define PSSTART start_subject  /* Field containing processed string start */
#define PSEND   end_subject    /* Field containing processed string end */

#include "pcre_internal.h"

/* Undefine some potentially clashing cpp symbols */

#undef min
#undef max

/* Values for setting in md->match_function_type to indicate two special types
of call to match(). We do it this way to save on using another stack variable,
as stack usage is to be discouraged. */

#define MATCH_CONDASSERT     1  /* Called to check a condition assertion */
#define MATCH_CBEGROUP       2  /* Could-be-empty unlimited repeat group */

/* Non-error returns from the match() function. Error returns are externally
defined PCRE_ERROR_xxx codes, which are all negative. */

#define MATCH_MATCH        1
#define MATCH_NOMATCH      0

/* Special internal returns from the match() function. Make them sufficiently
negative to avoid the external error codes. */

#define MATCH_ACCEPT       (-999)
#define MATCH_COMMIT       (-998)
#define MATCH_KETRPOS      (-997)
#define MATCH_ONCE         (-996)
#define MATCH_PRUNE        (-995)
#define MATCH_SKIP         (-994)
#define MATCH_SKIP_ARG     (-993)
#define MATCH_THEN         (-992)

/* This is a convenience macro for code that occurs many times. */

#define MRRETURN(ra) \
  { \
  md->mark = markptr; \
  RRETURN(ra); \
  }

/* Maximum number of ints of offset to save on the stack for recursive calls.
If the offset vector is bigger, malloc is used. This should be a multiple of 3,
because the offset vector is always a multiple of 3 long. */

#define REC_STACK_SAVE_MAX 30

/* Min and max values for the common repeats; for the maxima, 0 => infinity */

static const char rep_min[] = { 0, 0, 1, 1, 0, 0 };
static const char rep_max[] = { 0, 0, 0, 0, 1, 1 };



#ifdef PCRE_DEBUG
/*************************************************
*        Debugging function to print chars       *
*************************************************/

/* Print a sequence of chars in printable format, stopping at the end of the
subject if the requested.

Arguments:
  p           points to characters
  length      number to print
  is_subject  TRUE if printing from within md->start_subject
  md          pointer to matching data block, if is_subject is TRUE

Returns:     nothing
*/

static void
pchars(const uschar *p, int length, BOOL is_subject, match_data *md)
{
unsigned int c;
if (is_subject && length > md->end_subject - p) length = md->end_subject - p;
while (length-- > 0)
  if (isprint(c = *(p++))) printf("%c", c); else printf("\\x%02x", c);
}
#endif



/*************************************************
*          Match a back-reference                *
*************************************************/

/* Normally, if a back reference hasn't been set, the length that is passed is
negative, so the match always fails. However, in JavaScript compatibility mode,
the length passed is zero. Note that in caseless UTF-8 mode, the number of
subject bytes matched may be different to the number of reference bytes.

Arguments:
  offset      index into the offset vector
  eptr        pointer into the subject
  length      length of reference to be matched (number of bytes)
  md          points to match data block
  caseless    TRUE if caseless

Returns:      < 0 if not matched, otherwise the number of subject bytes matched
*/

static int
match_ref(int offset, register USPTR eptr, int length, match_data *md,
  BOOL caseless)
{
USPTR eptr_start = eptr;
register USPTR p = md->start_subject + md->offset_vector[offset];

#ifdef PCRE_DEBUG
if (eptr >= md->end_subject)
  printf("matching subject <null>");
else
  {
  printf("matching subject ");
  pchars(eptr, length, TRUE, md);
  }
printf(" against backref ");
pchars(p, length, FALSE, md);
printf("\n");
#endif

/* Always fail if reference not set (and not JavaScript compatible). */

if (length < 0) return -1;

/* Separate the caseless case for speed. In UTF-8 mode we can only do this
properly if Unicode properties are supported. Otherwise, we can check only
ASCII characters. */

if (caseless)
  {
#ifdef SUPPORT_UTF8
#ifdef SUPPORT_UCP
  if (md->utf8)
    {
    /* Match characters up to the end of the reference. NOTE: the number of
    bytes matched may differ, because there are some characters whose upper and
    lower case versions code as different numbers of bytes. For example, U+023A
    (2 bytes in UTF-8) is the upper case version of U+2C65 (3 bytes in UTF-8);
    a sequence of 3 of the former uses 6 bytes, as does a sequence of two of
    the latter. It is important, therefore, to check the length along the
    reference, not along the subject (earlier code did this wrong). */

    USPTR endptr = p + length;
    while (p < endptr)
      {
      int c, d;
      if (eptr >= md->end_subject) return -1;
      TGETCHARINC(c, eptr);
      TGETCHARINC(d, p);
      if (c != d && c != UCD_OTHERCASE(d)) return -1;
      }
    }
  else
#endif
#endif

  /* The same code works when not in UTF-8 mode and in UTF-8 mode when there
  is no UCP support. */
    {
    if (eptr + length > md->end_subject) return -1;
    while (length-- > 0)
      { if (md->lcc[*p++] != md->lcc[*eptr++]) return -1; }
    }
  }

/* In the caseful case, we can just compare the bytes, whether or not we
are in UTF-8 mode. */

else
  {
  if (eptr + length > md->end_subject) return -1;
  while (length-- > 0) if (*p++ != *eptr++) return -1;
  }

return (int)(eptr - eptr_start);
}



/***************************************************************************
****************************************************************************
                   RECURSION IN THE match() FUNCTION

The match() function is highly recursive, though not every recursive call
increases the recursive depth. Nevertheless, some regular expressions can cause
it to recurse to a great depth. I was writing for Unix, so I just let it call
itself recursively. This uses the stack for saving everything that has to be
saved for a recursive call. On Unix, the stack can be large, and this works
fine.

It turns out that on some non-Unix-like systems there are problems with
programs that use a lot of stack. (This despite the fact that every last chip
has oodles of memory these days, and techniques for extending the stack have
been known for decades.) So....

There is a fudge, triggered by defining NO_RECURSE, which avoids recursive
calls by keeping local variables that need to be preserved in blocks of memory
obtained from malloc() instead instead of on the stack. Macros are used to
achieve this so that the actual code doesn't look very different to what it
always used to.

The original heap-recursive code used longjmp(). However, it seems that this
can be very slow on some operating systems. Following a suggestion from Stan
Switzer, the use of longjmp() has been abolished, at the cost of having to
provide a unique number for each call to RMATCH. There is no way of generating
a sequence of numbers at compile time in C. I have given them names, to make
them stand out more clearly.

Crude tests on x86 Linux show a small speedup of around 5-8%. However, on
FreeBSD, avoiding longjmp() more than halves the time taken to run the standard
tests. Furthermore, not using longjmp() means that local dynamic variables
don't have indeterminate values; this has meant that the frame size can be
reduced because the result can be "passed back" by straight setting of the
variable instead of being passed in the frame.
****************************************************************************
***************************************************************************/

/* Numbers for RMATCH calls. When this list is changed, the code at HEAP_RETURN
below must be updated in sync.  */

enum { RM1=1, RM2,  RM3,  RM4,  RM5,  RM6,  RM7,  RM8,  RM9,  RM10,
       RM11,  RM12, RM13, RM14, RM15, RM16, RM17, RM18, RM19, RM20,
       RM21,  RM22, RM23, RM24, RM25, RM26, RM27, RM28, RM29, RM30,
       RM31,  RM32, RM33, RM34, RM35, RM36, RM37, RM38, RM39, RM40,
       RM41,  RM42, RM43, RM44, RM45, RM46, RM47, RM48, RM49, RM50,
       RM51,  RM52, RM53, RM54, RM55, RM56, RM57, RM58, RM59, RM60,
       RM61,  RM62, RM63 };

/* These versions of the macros use the stack, as normal. There are debugging
versions and production versions. Note that the "rw" argument of RMATCH isn't
actually used in this definition. */

#ifndef NO_RECURSE
#define REGISTER register

#ifdef PCRE_DEBUG
#define RMATCH(ra,rb,rc,rd,re,rw) \
  { \
  printf("match() called in line %d\n", __LINE__); \
  rrc = match(ra,rb,mstart,markptr,rc,rd,re,rdepth+1); \
  printf("to line %d\n", __LINE__); \
  }
#define RRETURN(ra) \
  { \
  printf("match() returned %d from line %d ", ra, __LINE__); \
  return ra; \
  }
#else
#define RMATCH(ra,rb,rc,rd,re,rw) \
  rrc = match(ra,rb,mstart,markptr,rc,rd,re,rdepth+1)
#define RRETURN(ra) return ra
#endif

#else


/* These versions of the macros manage a private stack on the heap. Note that
the "rd" argument of RMATCH isn't actually used in this definition. It's the md
argument of match(), which never changes. */

#define REGISTER

#define RMATCH(ra,rb,rc,rd,re,rw)\
  {\
  heapframe *newframe = (heapframe *)(pcre_stack_malloc)(sizeof(heapframe));\
  if (newframe == NULL) RRETURN(PCRE_ERROR_NOMEMORY);\
  frame->Xwhere = rw; \
  newframe->Xeptr = ra;\
  newframe->Xecode = rb;\
  newframe->Xmstart = mstart;\
  newframe->Xmarkptr = markptr;\
  newframe->Xoffset_top = rc;\
  newframe->Xeptrb = re;\
  newframe->Xrdepth = frame->Xrdepth + 1;\
  newframe->Xprevframe = frame;\
  frame = newframe;\
  DPRINTF(("restarting from line %d\n", __LINE__));\
  goto HEAP_RECURSE;\
  L_##rw:\
  DPRINTF(("jumped back to line %d\n", __LINE__));\
  }

#define RRETURN(ra)\
  {\
  heapframe *oldframe = frame;\
  frame = oldframe->Xprevframe;\
  (pcre_stack_free)(oldframe);\
  if (frame != NULL)\
    {\
    rrc = ra;\
    goto HEAP_RETURN;\
    }\
  return ra;\
  }


/* Structure for remembering the local variables in a private frame */

typedef struct heapframe {
  struct heapframe *Xprevframe;

  /* Function arguments that may change */

  USPTR Xeptr;
  const uschar *Xecode;
  USPTR Xmstart;
  USPTR Xmarkptr;
  int Xoffset_top;
  eptrblock *Xeptrb;
  unsigned int Xrdepth;

  /* Function local variables */

  USPTR Xcallpat;
#ifdef SUPPORT_UTF8
  USPTR Xcharptr;
#endif
  USPTR Xdata;
  USPTR Xnext;
  USPTR Xpp;
  USPTR Xprev;
  USPTR Xsaved_eptr;

  recursion_info Xnew_recursive;

  BOOL Xcur_is_word;
  BOOL Xcondition;
  BOOL Xprev_is_word;

#ifdef SUPPORT_UCP
  int Xprop_type;
  int Xprop_value;
  int Xprop_fail_result;
  int Xoclength;
  uschar Xocchars[8];
#endif

  int Xcodelink;
  int Xctype;
  unsigned int Xfc;
  int Xfi;
  int Xlength;
  int Xmax;
  int Xmin;
  int Xnumber;
  int Xoffset;
  int Xop;
  int Xsave_capture_last;
  int Xsave_offset1, Xsave_offset2, Xsave_offset3;
  int Xstacksave[REC_STACK_SAVE_MAX];

  eptrblock Xnewptrb;

  /* Where to jump back to */

  int Xwhere;

} heapframe;

#endif


/***************************************************************************
***************************************************************************/



/*************************************************
*         Match from current position            *
*************************************************/

/* This function is called recursively in many circumstances. Whenever it
returns a negative (error) response, the outer incarnation must also return the
same response. */

/* These macros pack up tests that are used for partial matching, and which
appears several times in the code. We set the "hit end" flag if the pointer is
at the end of the subject and also past the start of the subject (i.e.
something has been matched). For hard partial matching, we then return
immediately. The second one is used when we already know we are past the end of
the subject. */

#define CHECK_PARTIAL()\
  if (md->partial != 0 && eptr >= md->end_subject && \
      eptr > md->start_used_ptr) \
    { \
    md->hitend = TRUE; \
    if (md->partial > 1) MRRETURN(PCRE_ERROR_PARTIAL); \
    }

#define SCHECK_PARTIAL()\
  if (md->partial != 0 && eptr > md->start_used_ptr) \
    { \
    md->hitend = TRUE; \
    if (md->partial > 1) MRRETURN(PCRE_ERROR_PARTIAL); \
    }


/* Performance note: It might be tempting to extract commonly used fields from
the md structure (e.g. utf8, end_subject) into individual variables to improve
performance. Tests using gcc on a SPARC disproved this; in the first case, it
made performance worse.

Arguments:
   eptr        pointer to current character in subject
   ecode       pointer to current position in compiled code
   mstart      pointer to the current match start position (can be modified
                 by encountering \K)
   markptr     pointer to the most recent MARK name, or NULL
   offset_top  current top pointer
   md          pointer to "static" info for the match
   eptrb       pointer to chain of blocks containing eptr at start of
                 brackets - for testing for empty matches
   rdepth      the recursion depth

Returns:       MATCH_MATCH if matched            )  these values are >= 0
               MATCH_NOMATCH if failed to match  )
               a negative MATCH_xxx value for PRUNE, SKIP, etc
               a negative PCRE_ERROR_xxx value if aborted by an error condition
                 (e.g. stopped by repeated call or recursion limit)
*/

static int
match(REGISTER USPTR eptr, REGISTER const uschar *ecode, USPTR mstart,
  const uschar *markptr, int offset_top, match_data *md, eptrblock *eptrb,
  unsigned int rdepth)
{
/* These variables do not need to be preserved over recursion in this function,
so they can be ordinary variables in all cases. Mark some of them with
"register" because they are used a lot in loops. */

register int  rrc;         /* Returns from recursive calls */
register int  i;           /* Used for loops not involving calls to RMATCH() */
register unsigned int c;   /* Character values not kept over RMATCH() calls */
#ifdef SUPPORT_UTF8 /* AutoHotkey: This helps detected unintended usages of utf8. */
#ifdef SUPPORT_UTF8_ONLY
    const BOOL utf8 = TRUE;
#else
    register BOOL utf8;        /* Local copy of UTF-8 flag for speed */
#endif
#endif /* AutoHotkey. */

BOOL minimize, possessive; /* Quantifier options */
BOOL caseless;
int condcode;

/* When recursion is not being used, all "local" variables that have to be
preserved over calls to RMATCH() are part of a "frame" which is obtained from
heap storage. Set up the top-level frame here; others are obtained from the
heap whenever RMATCH() does a "recursion". See the macro definitions above. */

#ifdef NO_RECURSE
heapframe *frame = (heapframe *)(pcre_stack_malloc)(sizeof(heapframe));
if (frame == NULL) RRETURN(PCRE_ERROR_NOMEMORY);
frame->Xprevframe = NULL;            /* Marks the top level */

/* Copy in the original argument variables */

frame->Xeptr = eptr;
frame->Xecode = ecode;
frame->Xmstart = mstart;
frame->Xmarkptr = markptr;
frame->Xoffset_top = offset_top;
frame->Xeptrb = eptrb;
frame->Xrdepth = rdepth;

/* This is where control jumps back to to effect "recursion" */

HEAP_RECURSE:

/* Macros make the argument variables come from the current frame */

#define eptr               frame->Xeptr
#define ecode              frame->Xecode
#define mstart             frame->Xmstart
#define markptr            frame->Xmarkptr
#define offset_top         frame->Xoffset_top
#define eptrb              frame->Xeptrb
#define rdepth             frame->Xrdepth

/* Ditto for the local variables */

#ifdef SUPPORT_UTF8
#define charptr            frame->Xcharptr
#endif
#define callpat            frame->Xcallpat
#define codelink           frame->Xcodelink
#define data               frame->Xdata
#define next               frame->Xnext
#define pp                 frame->Xpp
#define prev               frame->Xprev
#define saved_eptr         frame->Xsaved_eptr

#define new_recursive      frame->Xnew_recursive

#define cur_is_word        frame->Xcur_is_word
#define condition          frame->Xcondition
#define prev_is_word       frame->Xprev_is_word

#ifdef SUPPORT_UCP
#define prop_type          frame->Xprop_type
#define prop_value         frame->Xprop_value
#define prop_fail_result   frame->Xprop_fail_result
#define oclength           frame->Xoclength
#define occhars            frame->Xocchars
#endif

#define ctype              frame->Xctype
#define fc                 frame->Xfc
#define fi                 frame->Xfi
#define length             frame->Xlength
#define max                frame->Xmax
#define min                frame->Xmin
#define number             frame->Xnumber
#define offset             frame->Xoffset
#define op                 frame->Xop
#define save_capture_last  frame->Xsave_capture_last
#define save_offset1       frame->Xsave_offset1
#define save_offset2       frame->Xsave_offset2
#define save_offset3       frame->Xsave_offset3
#define stacksave          frame->Xstacksave

#define newptrb            frame->Xnewptrb

/* When recursion is being used, local variables are allocated on the stack and
get preserved during recursion in the normal way. In this environment, fi and
i, and fc and c, can be the same variables. */

#else         /* NO_RECURSE not defined */
#define fi i
#define fc c

/* Many of the following variables are used only in small blocks of the code.
My normal style of coding would have declared them within each of those blocks.
However, in order to accommodate the version of this code that uses an external
"stack" implemented on the heap, it is easier to declare them all here, so the
declarations can be cut out in a block. The only declarations within blocks
below are for variables that do not have to be preserved over a recursive call
to RMATCH(). */

#ifdef SUPPORT_UTF8
const utchar *charptr;
#endif
const uschar *callpat;
const uschar *data;
const uschar *next;
USPTR         pp;
const uschar *prev;
USPTR         saved_eptr;

recursion_info new_recursive;

BOOL cur_is_word;
BOOL condition;
BOOL prev_is_word;

#ifdef SUPPORT_UCP
int prop_type;
int prop_value;
int prop_fail_result;
int oclength;
utchar occhars[8];
#endif

int codelink;
int ctype;
int length;
int max;
int min;
int number;
int offset;
int op;
int save_capture_last;
int save_offset1, save_offset2, save_offset3;
int stacksave[REC_STACK_SAVE_MAX];

eptrblock newptrb;
#endif     /* NO_RECURSE */

/* To save space on the stack and in the heap frame, I have doubled up on some
of the local variables that are used only in localised parts of the code, but
still need to be preserved over recursive calls of match(). These macros define
the alternative names that are used. */

#define allow_zero    cur_is_word
#define cbegroup      condition
#define code_offset   codelink
#define condassert    condition
#define matched_once  prev_is_word

/* These statements are here to stop the compiler complaining about unitialized
variables. */

#ifdef SUPPORT_UCP
prop_value = 0;
prop_fail_result = 0;
#endif


/* This label is used for tail recursion, which is used in a few cases even
when NO_RECURSE is not defined, in order to reduce the amount of stack that is
used. Thanks to Ian Taylor for noticing this possibility and sending the
original patch. */

TAIL_RECURSE:

/* OK, now we can get on with the real code of the function. Recursive calls
are specified by the macro RMATCH and RRETURN is used to return. When
NO_RECURSE is *not* defined, these just turn into a recursive call to match()
and a "return", respectively (possibly with some debugging if PCRE_DEBUG is
defined). However, RMATCH isn't like a function call because it's quite a
complicated macro. It has to be used in one particular way. This shouldn't,
however, impact performance when true recursion is being used. */

#if defined SUPPORT_UTF8 && !defined SUPPORT_UTF8_ONLY
utf8 = md->utf8;       /* Local copy of the flag */
/* AutoHotkey: Commented out to help detect unintended usages of utf8:
#else
utf8 = FALSE; */
#endif

/* First check that we haven't called match() too many times, or that we
haven't exceeded the recursive call limit. */

if (md->match_call_count++ >= md->match_limit) RRETURN(PCRE_ERROR_MATCHLIMIT);
if (rdepth >= md->match_limit_recursion) RRETURN(PCRE_ERROR_RECURSIONLIMIT);

/* At the start of a group with an unlimited repeat that may match an empty
string, the variable md->match_function_type is set to MATCH_CBEGROUP. It is
done this way to save having to use another function argument, which would take
up space on the stack. See also MATCH_CONDASSERT below.

When MATCH_CBEGROUP is set, add the current subject pointer to the chain of
such remembered pointers, to be checked when we hit the closing ket, in order
to break infinite loops that match no characters. When match() is called in
other circumstances, don't add to the chain. The MATCH_CBEGROUP feature must
NOT be used with tail recursion, because the memory block that is used is on
the stack, so a new one may be required for each match(). */

if (md->match_function_type == MATCH_CBEGROUP)
  {
  newptrb.epb_saved_eptr = eptr;
  newptrb.epb_prev = eptrb;
  eptrb = &newptrb;
  md->match_function_type = 0;
  }

/* Now start processing the opcodes. */

for (;;)
  {
  minimize = possessive = FALSE;
  op = *ecode;

  switch(op)
    {
    case OP_MARK:
    markptr = ecode + 2;
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode] + ecode[1], offset_top, md,
      eptrb, RM55);

    /* A return of MATCH_SKIP_ARG means that matching failed at SKIP with an
    argument, and we must check whether that argument matches this MARK's
    argument. It is passed back in md->start_match_ptr (an overloading of that
    variable). If it does match, we reset that variable to the current subject
    position and return MATCH_SKIP. Otherwise, pass back the return code
    unaltered. */

    if (rrc == MATCH_SKIP_ARG &&
        strcmp((char *)markptr, (char *)(md->start_match_ptr)) == 0)
      {
      md->start_match_ptr = eptr;
      RRETURN(MATCH_SKIP);
      }

    if (md->mark == NULL) md->mark = markptr;
    RRETURN(rrc);

    case OP_FAIL:
    MRRETURN(MATCH_NOMATCH);

    /* COMMIT overrides PRUNE, SKIP, and THEN */

    case OP_COMMIT:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
      eptrb, RM52);
    if (rrc != MATCH_NOMATCH && rrc != MATCH_PRUNE &&
        rrc != MATCH_SKIP && rrc != MATCH_SKIP_ARG &&
        rrc != MATCH_THEN)
      RRETURN(rrc);
    MRRETURN(MATCH_COMMIT);

    /* PRUNE overrides THEN */

    case OP_PRUNE:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
      eptrb, RM51);
    if (rrc != MATCH_NOMATCH && rrc != MATCH_THEN) RRETURN(rrc);
    MRRETURN(MATCH_PRUNE);

    case OP_PRUNE_ARG:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode] + ecode[1], offset_top, md,
      eptrb, RM56);
    if (rrc != MATCH_NOMATCH && rrc != MATCH_THEN) RRETURN(rrc);
    md->mark = ecode + 2;
    RRETURN(MATCH_PRUNE);

    /* SKIP overrides PRUNE and THEN */

    case OP_SKIP:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
      eptrb, RM53);
    if (rrc != MATCH_NOMATCH && rrc != MATCH_PRUNE && rrc != MATCH_THEN)
      RRETURN(rrc);
    md->start_match_ptr = eptr;   /* Pass back current position */
    MRRETURN(MATCH_SKIP);

    case OP_SKIP_ARG:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode] + ecode[1], offset_top, md,
      eptrb, RM57);
    if (rrc != MATCH_NOMATCH && rrc != MATCH_PRUNE && rrc != MATCH_THEN)
      RRETURN(rrc);

    /* Pass back the current skip name by overloading md->start_match_ptr and
    returning the special MATCH_SKIP_ARG return code. This will either be
    caught by a matching MARK, or get to the top, where it is treated the same
    as PRUNE. */

    md->start_match_ptr = (USPTR)(ecode + 2);
    RRETURN(MATCH_SKIP_ARG);

    /* For THEN (and THEN_ARG) we pass back the address of the bracket or
    the alt that is at the start of the current branch. This makes it possible
    to skip back past alternatives that precede the THEN within the current
    branch. */

    case OP_THEN:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
      eptrb, RM54);
    if (rrc != MATCH_NOMATCH) RRETURN(rrc);
    md->start_match_ptr = (USPTR)(ecode - GET(ecode, 1));
    MRRETURN(MATCH_THEN);

    case OP_THEN_ARG:
    RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode] + ecode[1+LINK_SIZE],
      offset_top, md, eptrb, RM58);
    if (rrc != MATCH_NOMATCH) RRETURN(rrc);
    md->start_match_ptr = (USPTR)(ecode - GET(ecode, 1));
    md->mark = ecode + LINK_SIZE + 2;
    RRETURN(MATCH_THEN);

    /* Handle a capturing bracket, other than those that are possessive with an
    unlimited repeat. If there is space in the offset vector, save the current
    subject position in the working slot at the top of the vector. We mustn't
    change the current values of the data slot, because they may be set from a
    previous iteration of this group, and be referred to by a reference inside
    the group. A failure to match might occur after the group has succeeded,
    if something later on doesn't match. For this reason, we need to restore
    the working value and also the values of the final offsets, in case they
    were set by a previous iteration of the same bracket.

    If there isn't enough space in the offset vector, treat this as if it were
    a non-capturing bracket. Don't worry about setting the flag for the error
    case here; that is handled in the code for KET. */

    case OP_CBRA:
    case OP_SCBRA:
    number = GET2(ecode, 1+LINK_SIZE);
    offset = number << 1;

#ifdef PCRE_DEBUG
    printf("start bracket %d\n", number);
    printf("subject=");
    pchars(eptr, 16, TRUE, md);
    printf("\n");
#endif

    if (offset < md->offset_max)
      {
      save_offset1 = md->offset_vector[offset];
      save_offset2 = md->offset_vector[offset+1];
      save_offset3 = md->offset_vector[md->offset_end - number];
      save_capture_last = md->capture_last;

      DPRINTF(("saving %d %d %d\n", save_offset1, save_offset2, save_offset3));
      md->offset_vector[md->offset_end - number] =
        (int)(eptr - md->start_subject);

      for (;;)
        {
        if (op >= OP_SBRA) md->match_function_type = MATCH_CBEGROUP;
        RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
          eptrb, RM1);
        if (rrc == MATCH_ONCE) break;  /* Backing up through an atomic group */
        if (rrc != MATCH_NOMATCH &&
            (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
          RRETURN(rrc);
        md->capture_last = save_capture_last;
        ecode += GET(ecode, 1);
        if (*ecode != OP_ALT) break;
        }

      DPRINTF(("bracket %d failed\n", number));
      md->offset_vector[offset] = save_offset1;
      md->offset_vector[offset+1] = save_offset2;
      md->offset_vector[md->offset_end - number] = save_offset3;

      /* At this point, rrc will be one of MATCH_ONCE, MATCH_NOMATCH, or
      MATCH_THEN. */

      if (rrc != MATCH_THEN && md->mark == NULL) md->mark = markptr;
      RRETURN(((rrc == MATCH_ONCE)? MATCH_ONCE:MATCH_NOMATCH));
      }

    /* FALL THROUGH ... Insufficient room for saving captured contents. Treat
    as a non-capturing bracket. */

    /* VVVVVVVVVVVVVVVVVVVVVVVVV */
    /* VVVVVVVVVVVVVVVVVVVVVVVVV */

    DPRINTF(("insufficient capture room: treat as non-capturing\n"));

    /* VVVVVVVVVVVVVVVVVVVVVVVVV */
    /* VVVVVVVVVVVVVVVVVVVVVVVVV */

    /* Non-capturing or atomic group, except for possessive with unlimited
    repeat. Loop for all the alternatives. When we get to the final alternative
    within the brackets, we used to return the result of a recursive call to
    match() whatever happened so it was possible to reduce stack usage by
    turning this into a tail recursion, except in the case of a possibly empty
    group. However, now that there is the possiblity of (*THEN) occurring in
    the final alternative, this optimization is no longer possible.

    MATCH_ONCE is returned when the end of an atomic group is successfully
    reached, but subsequent matching fails. It passes back up the tree (causing
    captured values to be reset) until the original atomic group level is
    reached. This is tested by comparing md->once_target with the start of the
    group. At this point, the return is converted into MATCH_NOMATCH so that
    previous backup points can be taken. */

    case OP_ONCE:
    case OP_BRA:
    case OP_SBRA:
    DPRINTF(("start non-capturing bracket\n"));

    for (;;)
      {
      if (op >= OP_SBRA || op == OP_ONCE) md->match_function_type = MATCH_CBEGROUP;
      RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md, eptrb,
        RM2);
      if (rrc != MATCH_NOMATCH &&
          (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
        {
        if (rrc == MATCH_ONCE)
          {
          const uschar *scode = ecode;
          if (*scode != OP_ONCE)           /* If not at start, find it */
            {
            while (*scode == OP_ALT) scode += GET(scode, 1);
            scode -= GET(scode, 1);
            }
          if (md->once_target == scode) rrc = MATCH_NOMATCH;
          }
        RRETURN(rrc);
        }
      ecode += GET(ecode, 1);
      if (*ecode != OP_ALT) break;
      }
    if (rrc != MATCH_THEN && md->mark == NULL) md->mark = markptr;
    RRETURN(MATCH_NOMATCH);

    /* Handle possessive capturing brackets with an unlimited repeat. We come
    here from BRAZERO with allow_zero set TRUE. The offset_vector values are
    handled similarly to the normal case above. However, the matching is
    different. The end of these brackets will always be OP_KETRPOS, which
    returns MATCH_KETRPOS without going further in the pattern. By this means
    we can handle the group by iteration rather than recursion, thereby
    reducing the amount of stack needed. */

    case OP_CBRAPOS:
    case OP_SCBRAPOS:
    allow_zero = FALSE;

    POSSESSIVE_CAPTURE:
    number = GET2(ecode, 1+LINK_SIZE);
    offset = number << 1;

#ifdef PCRE_DEBUG
    printf("start possessive bracket %d\n", number);
    printf("subject=");
    pchars(eptr, 16, TRUE, md);
    printf("\n");
#endif

    if (offset < md->offset_max)
      {
      matched_once = FALSE;
      code_offset = (int)(ecode - md->start_code);

      save_offset1 = md->offset_vector[offset];
      save_offset2 = md->offset_vector[offset+1];
      save_offset3 = md->offset_vector[md->offset_end - number];
      save_capture_last = md->capture_last;

      DPRINTF(("saving %d %d %d\n", save_offset1, save_offset2, save_offset3));

      /* Each time round the loop, save the current subject position for use
      when the group matches. For MATCH_MATCH, the group has matched, so we
      restart it with a new subject starting position, remembering that we had
      at least one match. For MATCH_NOMATCH, carry on with the alternatives, as
      usual. If we haven't matched any alternatives in any iteration, check to
      see if a previous iteration matched. If so, the group has matched;
      continue from afterwards. Otherwise it has failed; restore the previous
      capture values before returning NOMATCH. */

      for (;;)
        {
        md->offset_vector[md->offset_end - number] =
          (int)(eptr - md->start_subject);
        if (op >= OP_SBRA) md->match_function_type = MATCH_CBEGROUP;
        RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
          eptrb, RM63);
        if (rrc == MATCH_KETRPOS)
          {
          offset_top = md->end_offset_top;
          eptr = md->end_match_ptr;
          ecode = md->start_code + code_offset;
          save_capture_last = md->capture_last;
          matched_once = TRUE;
          continue;
          }
        if (rrc != MATCH_NOMATCH &&
            (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
          RRETURN(rrc);
        md->capture_last = save_capture_last;
        ecode += GET(ecode, 1);
        if (*ecode != OP_ALT) break;
        }

      if (!matched_once)
        {
        md->offset_vector[offset] = save_offset1;
        md->offset_vector[offset+1] = save_offset2;
        md->offset_vector[md->offset_end - number] = save_offset3;
        }

      if (rrc != MATCH_THEN && md->mark == NULL) md->mark = markptr;
      if (allow_zero || matched_once)
        {
        ecode += 1 + LINK_SIZE;
        break;
        }

      RRETURN(MATCH_NOMATCH);
      }

    /* FALL THROUGH ... Insufficient room for saving captured contents. Treat
    as a non-capturing bracket. */

    /* VVVVVVVVVVVVVVVVVVVVVVVVV */
    /* VVVVVVVVVVVVVVVVVVVVVVVVV */

    DPRINTF(("insufficient capture room: treat as non-capturing\n"));

    /* VVVVVVVVVVVVVVVVVVVVVVVVV */
    /* VVVVVVVVVVVVVVVVVVVVVVVVV */

    /* Non-capturing possessive bracket with unlimited repeat. We come here
    from BRAZERO with allow_zero = TRUE. The code is similar to the above,
    without the capturing complication. It is written out separately for speed
    and cleanliness. */

    case OP_BRAPOS:
    case OP_SBRAPOS:
    allow_zero = FALSE;

    POSSESSIVE_NON_CAPTURE:
    matched_once = FALSE;
    code_offset = (int)(ecode - md->start_code);

    for (;;)
      {
      if (op >= OP_SBRA) md->match_function_type = MATCH_CBEGROUP;
      RMATCH(eptr, ecode + _pcre_OP_lengths[*ecode], offset_top, md,
        eptrb, RM48);
      if (rrc == MATCH_KETRPOS)
        {
        offset_top = md->end_offset_top;
        eptr = md->end_match_ptr;
        ecode = md->start_code + code_offset;
        matched_once = TRUE;
        continue;
        }
      if (rrc != MATCH_NOMATCH &&
          (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
        RRETURN(rrc);
      ecode += GET(ecode, 1);
      if (*ecode != OP_ALT) break;
      }

    if (matched_once || allow_zero)
      {
      ecode += 1 + LINK_SIZE;
      break;
      }
    RRETURN(MATCH_NOMATCH);

    /* Control never reaches here. */

    /* Conditional group: compilation checked that there are no more than
    two branches. If the condition is false, skipping the first branch takes us
    past the end if there is only one branch, but that's OK because that is
    exactly what going to the ket would do. */

    case OP_COND:
    case OP_SCOND:
    codelink = GET(ecode, 1);

    /* Because of the way auto-callout works during compile, a callout item is
    inserted between OP_COND and an assertion condition. */

    if (ecode[LINK_SIZE+1] == OP_CALLOUT)
      {
      if (pcre_callout != NULL)
        {
        pcre_callout_block cb;
        cb.version          = 2;   /* Version 1 of the callout block */
        cb.callout_number   = ecode[LINK_SIZE+2];
        cb.offset_vector    = md->offset_vector;
        cb.subject          = (PCRE_SPTR)md->start_subject;
        cb.subject_length   = (int)(md->end_subject - md->start_subject);
        cb.start_match      = (int)(mstart - md->start_subject);
        cb.current_position = (int)(eptr - md->start_subject);
        cb.pattern_position = GET(ecode, LINK_SIZE + 3);
        cb.next_item_length = GET(ecode, 3 + 2*LINK_SIZE);
        cb.capture_top      = offset_top/2;
        cb.capture_last     = md->capture_last;
        cb.callout_data     = md->callout_data;
        cb.mark             = markptr;
        if ((rrc = (*pcre_callout)(&cb)) > 0) MRRETURN(MATCH_NOMATCH);
        if (rrc < 0) RRETURN(rrc);
        }
      ecode += _pcre_OP_lengths[OP_CALLOUT];
      }

    condcode = ecode[LINK_SIZE+1];

    /* Now see what the actual condition is */

    if (condcode == OP_RREF || condcode == OP_NRREF)    /* Recursion test */
      {
      if (md->recursive == NULL)                /* Not recursing => FALSE */
        {
        condition = FALSE;
        ecode += GET(ecode, 1);
        }
      else
        {
        int recno = GET2(ecode, LINK_SIZE + 2);   /* Recursion group number*/
        condition =  (recno == RREF_ANY || recno == md->recursive->group_num);

        /* If the test is for recursion into a specific subpattern, and it is
        false, but the test was set up by name, scan the table to see if the
        name refers to any other numbers, and test them. The condition is true
        if any one is set. */

        if (!condition && condcode == OP_NRREF && recno != RREF_ANY)
          {
          utchar *slotA = md->name_table;
          for (i = 0; i < md->name_count; i++)
            {
            if (GET2(slotA, 0) == recno) break;
            slotA += md->name_entry_size;
            }

          /* Found a name for the number - there can be only one; duplicate
          names for different numbers are allowed, but not vice versa. First
          scan down for duplicates. */

          if (i < md->name_count)
            {
            utchar *slotB = slotA;
            while (slotB > md->name_table)
              {
              slotB -= md->name_entry_size;
              if (tstrcmp(slotA + 2, slotB + 2) == 0)
                {
                condition = GET2(slotB, 0) == md->recursive->group_num;
                if (condition) break;
                }
              else break;
              }

            /* Scan up for duplicates */

            if (!condition)
              {
              slotB = slotA;
              for (i++; i < md->name_count; i++)
                {
                slotB += md->name_entry_size;
                if (tstrcmp(slotA + 2, slotB + 2) == 0)
                  {
                  condition = GET2(slotB, 0) == md->recursive->group_num;
                  if (condition) break;
                  }
                else break;
                }
              }
            }
          }

        /* Chose branch according to the condition */

        ecode += condition? 3 : GET(ecode, 1);
        }
      }

    else if (condcode == OP_CREF || condcode == OP_NCREF)  /* Group used test */
      {
      offset = GET2(ecode, LINK_SIZE+2) << 1;  /* Doubled ref number */
      condition = offset < offset_top && md->offset_vector[offset] >= 0;

      /* If the numbered capture is unset, but the reference was by name,
      scan the table to see if the name refers to any other numbers, and test
      them. The condition is true if any one is set. This is tediously similar
      to the code above, but not close enough to try to amalgamate. */

      if (!condition && condcode == OP_NCREF)
        {
        int refno = offset >> 1;
        utchar *slotA = md->name_table;

        for (i = 0; i < md->name_count; i++)
          {
          if (GET2(slotA, 0) == refno) break;
          slotA += md->name_entry_size;
          }

        /* Found a name for the number - there can be only one; duplicate names
        for different numbers are allowed, but not vice versa. First scan down
        for duplicates. */

        if (i < md->name_count)
          {
          utchar *slotB = slotA;
          while (slotB > md->name_table)
            {
            slotB -= md->name_entry_size;
            if (tstrcmp(slotA + 2, slotB + 2) == 0)
              {
              offset = GET2(slotB, 0) << 1;
              condition = offset < offset_top &&
                md->offset_vector[offset] >= 0;
              if (condition) break;
              }
            else break;
            }

          /* Scan up for duplicates */

          if (!condition)
            {
            slotB = slotA;
            for (i++; i < md->name_count; i++)
              {
              slotB += md->name_entry_size;
              if (tstrcmp(slotA + 2, slotB + 2) == 0)
                {
                offset = GET2(slotB, 0) << 1;
                condition = offset < offset_top &&
                  md->offset_vector[offset] >= 0;
                if (condition) break;
                }
              else break;
              }
            }
          }
        }

      /* Chose branch according to the condition */

      ecode += condition? 3 : GET(ecode, 1);
      }

    else if (condcode == OP_DEF)     /* DEFINE - always false */
      {
      condition = FALSE;
      ecode += GET(ecode, 1);
      }

    /* The condition is an assertion. Call match() to evaluate it - setting
    md->match_function_type to MATCH_CONDASSERT causes it to stop at the end of
    an assertion. */

    else
      {
      md->match_function_type = MATCH_CONDASSERT;
      RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, NULL, RM3);
      if (rrc == MATCH_MATCH)
        {
        if (md->end_offset_top > offset_top)
          offset_top = md->end_offset_top;  /* Captures may have happened */
        condition = TRUE;
        ecode += 1 + LINK_SIZE + GET(ecode, LINK_SIZE + 2);
        while (*ecode == OP_ALT) ecode += GET(ecode, 1);
        }
      else if (rrc != MATCH_NOMATCH &&
              (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
        {
        RRETURN(rrc);         /* Need braces because of following else */
        }
      else
        {
        condition = FALSE;
        ecode += codelink;
        }
      }

    /* We are now at the branch that is to be obeyed. As there is only one,
    we used to use tail recursion to avoid using another stack frame, except
    when there was unlimited repeat of a possibly empty group. However, that
    strategy no longer works because of the possibilty of (*THEN) being
    encountered in the branch. A recursive call to match() is always required,
    unless the second alternative doesn't exist, in which case we can just
    plough on. */

    if (condition || *ecode == OP_ALT)
      {
      if (op == OP_SCOND) md->match_function_type = MATCH_CBEGROUP;
      RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, eptrb, RM49);
      if (rrc == MATCH_THEN && md->start_match_ptr == (USPTR)ecode)
        rrc = MATCH_NOMATCH;
      RRETURN(rrc);
      }
    else                         /* Condition false & no alternative */
      {
      ecode += 1 + LINK_SIZE;
      }
    break;


    /* Before OP_ACCEPT there may be any number of OP_CLOSE opcodes,
    to close any currently open capturing brackets. */

    case OP_CLOSE:
    number = GET2(ecode, 1);
    offset = number << 1;

#ifdef PCRE_DEBUG
      printf("end bracket %d at *ACCEPT", number);
      printf("\n");
#endif

    md->capture_last = number;
    if (offset >= md->offset_max) md->offset_overflow = TRUE; else
      {
      md->offset_vector[offset] =
        md->offset_vector[md->offset_end - number];
      md->offset_vector[offset+1] = (int)(eptr - md->start_subject);
      if (offset_top <= offset) offset_top = offset + 2;
      }
    ecode += 3;
    break;


    /* End of the pattern, either real or forced. */

    case OP_END:
    case OP_ACCEPT:
    case OP_ASSERT_ACCEPT:

    /* If we have matched an empty string, fail if not in an assertion and not
    in a recursion if either PCRE_NOTEMPTY is set, or if PCRE_NOTEMPTY_ATSTART
    is set and we have matched at the start of the subject. In both cases,
    backtracking will then try other alternatives, if any. */

    if (eptr == mstart && op != OP_ASSERT_ACCEPT &&
         md->recursive == NULL &&
         (md->notempty ||
           (md->notempty_atstart &&
             mstart == md->start_subject + md->start_offset)))
      MRRETURN(MATCH_NOMATCH);

    /* Otherwise, we have a match. */

    md->end_match_ptr = eptr;           /* Record where we ended */
    md->end_offset_top = offset_top;    /* and how many extracts were taken */
    md->start_match_ptr = mstart;       /* and the start (\K can modify) */

    /* For some reason, the macros don't work properly if an expression is
    given as the argument to MRRETURN when the heap is in use. */

    rrc = (op == OP_END)? MATCH_MATCH : MATCH_ACCEPT;
    MRRETURN(rrc);

    /* Assertion brackets. Check the alternative branches in turn - the
    matching won't pass the KET for an assertion. If any one branch matches,
    the assertion is true. Lookbehind assertions have an OP_REVERSE item at the
    start of each branch to move the current point backwards, so the code at
    this level is identical to the lookahead case. When the assertion is part
    of a condition, we want to return immediately afterwards. The caller of
    this incarnation of the match() function will have set MATCH_CONDASSERT in
    md->match_function type, and one of these opcodes will be the first opcode
    that is processed. We use a local variable that is preserved over calls to
    match() to remember this case. */

    case OP_ASSERT:
    case OP_ASSERTBACK:
    if (md->match_function_type == MATCH_CONDASSERT)
      {
      condassert = TRUE;
      md->match_function_type = 0;
      }
    else condassert = FALSE;

    do
      {
      RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, NULL, RM4);
      if (rrc == MATCH_MATCH || rrc == MATCH_ACCEPT)
        {
        mstart = md->start_match_ptr;   /* In case \K reset it */
        markptr = md->mark;
        break;
        }
      if (rrc != MATCH_NOMATCH &&
          (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
        RRETURN(rrc);
      ecode += GET(ecode, 1);
      }
    while (*ecode == OP_ALT);

    if (*ecode == OP_KET) MRRETURN(MATCH_NOMATCH);

    /* If checking an assertion for a condition, return MATCH_MATCH. */

    if (condassert) RRETURN(MATCH_MATCH);

    /* Continue from after the assertion, updating the offsets high water
    mark, since extracts may have been taken during the assertion. */

    do ecode += GET(ecode,1); while (*ecode == OP_ALT);
    ecode += 1 + LINK_SIZE;
    offset_top = md->end_offset_top;
    continue;

    /* Negative assertion: all branches must fail to match. Encountering SKIP,
    PRUNE, or COMMIT means we must assume failure without checking subsequent
    branches. */

    case OP_ASSERT_NOT:
    case OP_ASSERTBACK_NOT:
    if (md->match_function_type == MATCH_CONDASSERT)
      {
      condassert = TRUE;
      md->match_function_type = 0;
      }
    else condassert = FALSE;

    do
      {
      RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, NULL, RM5);
      if (rrc == MATCH_MATCH || rrc == MATCH_ACCEPT) MRRETURN(MATCH_NOMATCH);
      if (rrc == MATCH_SKIP || rrc == MATCH_PRUNE || rrc == MATCH_COMMIT)
        {
        do ecode += GET(ecode,1); while (*ecode == OP_ALT);
        break;
        }
      if (rrc != MATCH_NOMATCH &&
          (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
        RRETURN(rrc);
      ecode += GET(ecode,1);
      }
    while (*ecode == OP_ALT);

    if (condassert) RRETURN(MATCH_MATCH);  /* Condition assertion */

    ecode += 1 + LINK_SIZE;
    continue;

    /* Move the subject pointer back. This occurs only at the start of
    each branch of a lookbehind assertion. If we are too close to the start to
    move back, this match function fails. When working with UTF-8 we move
    back a number of characters, not bytes. */

    case OP_REVERSE:
#ifdef SUPPORT_UTF8
    if (utf8)
      {
      i = GET(ecode, 1);
      while (i-- > 0)
        {
        eptr--;
        if (eptr < md->start_subject) MRRETURN(MATCH_NOMATCH);
        TBACKCHAR(eptr);
        }
      }
    else
#endif

    /* No UTF-8 support, or not in UTF-8 mode: count is byte count */

      {
      eptr -= GET(ecode, 1);
      if (eptr < md->start_subject) MRRETURN(MATCH_NOMATCH);
      }

    /* Save the earliest consulted character, then skip to next op code */

    if (eptr < md->start_used_ptr) md->start_used_ptr = eptr;
    ecode += 1 + LINK_SIZE;
    break;

    /* The callout item calls an external function, if one is provided, passing
    details of the match so far. This is mainly for debugging, though the
    function is able to force a failure. */

    case OP_CALLOUT:
#ifdef SUPPORT_CALLOUT  /* AutoHotkey: Omit the callout feature from the code until it's needed. */
    if (pcre_callout != NULL)
      {
      pcre_callout_block cb;
      cb.version          = 2;   /* Version 1 of the callout block */
      cb.callout_number   = ecode[1];
      cb.offset_vector    = md->offset_vector;
      cb.subject          = (PCRE_SPTR)md->start_subject;
      cb.subject_length   = (int)(md->end_subject - md->start_subject);
      cb.start_match      = (int)(mstart - md->start_subject);
      cb.current_position = (int)(eptr - md->start_subject);
      cb.pattern_position = GET(ecode, 2 + sizeof(void **));
      cb.next_item_length = GET(ecode, 2 + sizeof(void **) + LINK_SIZE);
      cb.capture_top      = offset_top/2;
      cb.capture_last     = md->capture_last;
      cb.callout_data     = md->callout_data;
      cb.mark             = markptr;
      cb.user_callout     = *(void **)(ecode + 2); /* AutoHotkey */
      if ((rrc = (*pcre_callout)(&cb)) > 0) MRRETURN(MATCH_NOMATCH);
      if (rrc < 0) RRETURN(rrc);
      }
#endif /* AutoHotkey */
    ecode += 2 + sizeof(void **) + 2*LINK_SIZE;
    break;

    /* Recursion either matches the current regex, or some subexpression. The
    offset data is the offset to the starting bracket from the start of the
    whole pattern. (This is so that it works from duplicated subpatterns.)

    The state of the capturing groups is preserved over recursion, and
    re-instated afterwards. We don't know how many are started and not yet
    finished (offset_top records the completed total) so we just have to save
    all the potential data. There may be up to 65535 such values, which is too
    large to put on the stack, but using malloc for small numbers seems
    expensive. As a compromise, the stack is used when there are no more than
    REC_STACK_SAVE_MAX values to store; otherwise malloc is used.

    There are also other values that have to be saved. We use a chained
    sequence of blocks that actually live on the stack. Thanks to Robin Houston
    for the original version of this logic. It has, however, been hacked around
    a lot, so he is not to blame for the current way it works. */

    case OP_RECURSE:
      {
      recursion_info *ri;
      int recno;

      callpat = md->start_code + GET(ecode, 1);
      recno = (callpat == md->start_code)? 0 :
        GET2(callpat, 1 + LINK_SIZE);

      /* Check for repeating a recursion without advancing the subject pointer.
      This should catch convoluted mutual recursions. (Some simple cases are
      caught at compile time.) */

      for (ri = md->recursive; ri != NULL; ri = ri->prevrec)
        if (recno == ri->group_num && eptr == ri->subject_position)
          RRETURN(PCRE_ERROR_RECURSELOOP);

      /* Add to "recursing stack" */

      new_recursive.group_num = recno;
      new_recursive.subject_position = eptr;
      new_recursive.prevrec = md->recursive;
      md->recursive = &new_recursive;

      /* Where to continue from afterwards */

      ecode += 1 + LINK_SIZE;

      /* Now save the offset data */

      new_recursive.saved_max = md->offset_end;
      if (new_recursive.saved_max <= REC_STACK_SAVE_MAX)
        new_recursive.offset_save = stacksave;
      else
        {
        new_recursive.offset_save =
          (int *)(pcre_malloc)(new_recursive.saved_max * sizeof(int));
        if (new_recursive.offset_save == NULL) RRETURN(PCRE_ERROR_NOMEMORY);
        }
      memcpy(new_recursive.offset_save, md->offset_vector,
            new_recursive.saved_max * sizeof(int));

      /* OK, now we can do the recursion. After processing each alternative,
      restore the offset data. If there were nested recursions, md->recursive
      might be changed, so reset it before looping. */

      DPRINTF(("Recursing into group %d\n", new_recursive.group_num));
      cbegroup = (*callpat >= OP_SBRA);
      do
        {
        if (cbegroup) md->match_function_type = MATCH_CBEGROUP;
        RMATCH(eptr, callpat + _pcre_OP_lengths[*callpat], offset_top,
          md, eptrb, RM6);
        memcpy(md->offset_vector, new_recursive.offset_save,
            new_recursive.saved_max * sizeof(int));
        if (rrc == MATCH_MATCH || rrc == MATCH_ACCEPT)
          {
          DPRINTF(("Recursion matched\n"));
          md->recursive = new_recursive.prevrec;
          if (new_recursive.offset_save != stacksave)
            (pcre_free)(new_recursive.offset_save);

          /* Set where we got to in the subject, and reset the start in case
          it was changed by \K. This *is* propagated back out of a recursion,
          for Perl compatibility. */

          eptr = md->end_match_ptr;
          mstart = md->start_match_ptr;
          goto RECURSION_MATCHED;        /* Exit loop; end processing */
          }
        else if (rrc != MATCH_NOMATCH &&
                (rrc != MATCH_THEN || md->start_match_ptr != (USPTR)ecode))
          {
          DPRINTF(("Recursion gave error %d\n", rrc));
          if (new_recursive.offset_save != stacksave)
            (pcre_free)(new_recursive.offset_save);
          RRETURN(rrc);
          }

        md->recursive = &new_recursive;
        callpat += GET(callpat, 1);
        }
      while (*callpat == OP_ALT);

      DPRINTF(("Recursion didn't match\n"));
      md->recursive = new_recursive.prevrec;
      if (new_recursive.offset_save != stacksave)
        (pcre_free)(new_recursive.offset_save);
      MRRETURN(MATCH_NOMATCH);
      }

    RECURSION_MATCHED:
    break;

    /* An alternation is the end of a branch; scan along to find the end of the
    bracketed group and go to there. */

    case OP_ALT:
    do ecode += GET(ecode,1); while (*ecode == OP_ALT);
    break;

    /* BRAZERO, BRAMINZERO and SKIPZERO occur just before a bracket group,
    indicating that it may occur zero times. It may repeat infinitely, or not
    at all - i.e. it could be ()* or ()? or even (){0} in the pattern. Brackets
    with fixed upper repeat limits are compiled as a number of copies, with the
    optional ones preceded by BRAZERO or BRAMINZERO. */

    case OP_BRAZERO:
    next = ecode + 1;
    RMATCH(eptr, next, offset_top, md, eptrb, RM10);
    if (rrc != MATCH_NOMATCH) RRETURN(rrc);
    do next += GET(next, 1); while (*next == OP_ALT);
    ecode = next + 1 + LINK_SIZE;
    break;

    case OP_BRAMINZERO:
    next = ecode + 1;
    do next += GET(next, 1); while (*next == OP_ALT);
    RMATCH(eptr, next + 1+LINK_SIZE, offset_top, md, eptrb, RM11);
    if (rrc != MATCH_NOMATCH) RRETURN(rrc);
    ecode++;
    break;

    case OP_SKIPZERO:
    next = ecode+1;
    do next += GET(next,1); while (*next == OP_ALT);
    ecode = next + 1 + LINK_SIZE;
    break;

    /* BRAPOSZERO occurs before a possessive bracket group. Don't do anything
    here; just jump to the group, with allow_zero set TRUE. */

    case OP_BRAPOSZERO:
    op = *(++ecode);
    allow_zero = TRUE;
    if (op == OP_CBRAPOS || op == OP_SCBRAPOS) goto POSSESSIVE_CAPTURE;
      goto POSSESSIVE_NON_CAPTURE;

    /* End of a group, repeated or non-repeating. */

    case OP_KET:
    case OP_KETRMIN:
    case OP_KETRMAX:
    case OP_KETRPOS:
    prev = ecode - GET(ecode, 1);

    /* If this was a group that remembered the subject start, in order to break
    infinite repeats of empty string matches, retrieve the subject start from
    the chain. Otherwise, set it NULL. */

    if (*prev >= OP_SBRA || *prev == OP_ONCE)
      {
      saved_eptr = eptrb->epb_saved_eptr;   /* Value at start of group */
      eptrb = eptrb->epb_prev;              /* Backup to previous group */
      }
    else saved_eptr = NULL;

    /* If we are at the end of an assertion group, stop matching and return
    MATCH_MATCH, but record the current high water mark for use by positive
    assertions. We also need to record the match start in case it was changed
    by \K. */

    if (*prev == OP_ASSERT || *prev == OP_ASSERT_NOT ||
        *prev == OP_ASSERTBACK || *prev == OP_ASSERTBACK_NOT)
      {
      md->end_match_ptr = eptr;      /* For ONCE */
      md->end_offset_top = offset_top;
      md->start_match_ptr = mstart;
      MRRETURN(MATCH_MATCH);         /* Sets md->mark */
      }

    /* For capturing groups we have to check the group number back at the start
    and if necessary complete handling an extraction by setting the offsets and
    bumping the high water mark. Whole-pattern recursion is coded as a recurse
    into group 0, so it won't be picked up here. Instead, we catch it when the
    OP_END is reached. Other recursion is handled here. We just have to record
    the current subject position and start match pointer and give a MATCH
    return. */

    if (*prev == OP_CBRA || *prev == OP_SCBRA ||
        *prev == OP_CBRAPOS || *prev == OP_SCBRAPOS)
      {
      number = GET2(prev, 1+LINK_SIZE);
      offset = number << 1;

#ifdef PCRE_DEBUG
      printf("end bracket %d", number);
      printf("\n");
#endif

      /* Handle a recursively called group. */

      if (md->recursive != NULL && md->recursive->group_num == number)
        {
        md->end_match_ptr = eptr;
        md->start_match_ptr = mstart;
        RRETURN(MATCH_MATCH);
        }

      /* Deal with capturing */

      md->capture_last = number;
      if (offset >= md->offset_max) md->offset_overflow = TRUE; else
        {
        /* If offset is greater than offset_top, it means that we are
        "skipping" a capturing group, and that group's offsets must be marked
        unset. In earlier versions of PCRE, all the offsets were unset at the
        start of matching, but this doesn't work because atomic groups and
        assertions can cause a value to be set that should later be unset.
        Example: matching /(?>(a))b|(a)c/ against "ac". This sets group 1 as
        part of the atomic group, but this is not on the final matching path,
        so must be unset when 2 is set. (If there is no group 2, there is no
        problem, because offset_top will then be 2, indicating no capture.) */

        if (offset > offset_top)
          {
          register int *iptr = md->offset_vector + offset_top;
          register int *iend = md->offset_vector + offset;
          while (iptr < iend) *iptr++ = -1;
          }

        /* Now make the extraction */

        md->offset_vector[offset] =
          md->offset_vector[md->offset_end - number];
        md->offset_vector[offset+1] = (int)(eptr - md->start_subject);
        if (offset_top <= offset) offset_top = offset + 2;
        }
      }

    /* For an ordinary non-repeating ket, just continue at this level. This
    also happens for a repeating ket if no characters were matched in the
    group. This is the forcible breaking of infinite loops as implemented in
    Perl 5.005. For a non-repeating atomic group, establish a backup point by
    processing the rest of the pattern at a lower level. If this results in a
    NOMATCH return, pass MATCH_ONCE back to the original OP_ONCE level, thereby
    bypassing intermediate backup points, but resetting any captures that
    happened along the way. */

    if (*ecode == OP_KET || eptr == saved_eptr)
      {
      if (*prev == OP_ONCE)
        {
        RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, eptrb, RM12);
        if (rrc != MATCH_NOMATCH) RRETURN(rrc);
        md->once_target = prev;  /* Level at which to change to MATCH_NOMATCH */
        RRETURN(MATCH_ONCE);
        }
      ecode += 1 + LINK_SIZE;    /* Carry on at this level */
      break;
      }

    /* OP_KETRPOS is a possessive repeating ket. Remember the current position,
    and return the MATCH_KETRPOS. This makes it possible to do the repeats one
    at a time from the outer level, thus saving stack. */

    if (*ecode == OP_KETRPOS)
      {
      md->end_match_ptr = eptr;
      md->end_offset_top = offset_top;
      RRETURN(MATCH_KETRPOS);
      }

    /* The normal repeating kets try the rest of the pattern or restart from
    the preceding bracket, in the appropriate order. In the second case, we can
    use tail recursion to avoid using another stack frame, unless we have an
    an atomic group or an unlimited repeat of a group that can match an empty
    string. */

    if (*ecode == OP_KETRMIN)
      {
      RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, eptrb, RM7);
      if (rrc != MATCH_NOMATCH) RRETURN(rrc);
      if (*prev == OP_ONCE)
        {
        RMATCH(eptr, prev, offset_top, md, eptrb, RM8);
        if (rrc != MATCH_NOMATCH) RRETURN(rrc);
        md->once_target = prev;  /* Level at which to change to MATCH_NOMATCH */
        RRETURN(MATCH_ONCE);
        }
      if (*prev >= OP_SBRA)    /* Could match an empty string */
        {
        md->match_function_type = MATCH_CBEGROUP;
        RMATCH(eptr, prev, offset_top, md, eptrb, RM50);
        RRETURN(rrc);
        }
      ecode = prev;
      goto TAIL_RECURSE;
      }
    else  /* OP_KETRMAX */
      {
      if (*prev >= OP_SBRA) md->match_function_type = MATCH_CBEGROUP;
      RMATCH(eptr, prev, offset_top, md, eptrb, RM13);
      if (rrc == MATCH_ONCE && md->once_target == prev) rrc = MATCH_NOMATCH;
      if (rrc != MATCH_NOMATCH) RRETURN(rrc);
      if (*prev == OP_ONCE)
        {
        RMATCH(eptr, ecode + 1 + LINK_SIZE, offset_top, md, eptrb, RM9);
        if (rrc != MATCH_NOMATCH) RRETURN(rrc);
        md->once_target = prev;
        RRETURN(MATCH_ONCE);
        }
      ecode += 1 + LINK_SIZE;
      goto TAIL_RECURSE;
      }
    /* Control never gets here */

    /* Not multiline mode: start of subject assertion, unless notbol. */

    case OP_CIRC:
    if (md->notbol && eptr == md->start_subject) MRRETURN(MATCH_NOMATCH);

    /* Start of subject assertion */

    case OP_SOD:
    if (eptr != md->start_subject) MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    /* Multiline mode: start of subject unless notbol, or after any newline. */

    case OP_CIRCM:
    if (md->notbol && eptr == md->start_subject) MRRETURN(MATCH_NOMATCH);
    if (eptr != md->start_subject &&
        (eptr == md->end_subject || !WAS_NEWLINE(eptr)))
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    /* Start of match assertion */

    case OP_SOM:
    if (eptr != md->start_subject + md->start_offset) MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    /* Reset the start of match point */

    case OP_SET_SOM:
    mstart = eptr;
    ecode++;
    break;

    /* Multiline mode: assert before any newline, or before end of subject
    unless noteol is set. */

    case OP_DOLLM:
    if (eptr < md->end_subject)
      { if (!IS_NEWLINE(eptr)) MRRETURN(MATCH_NOMATCH); }
    else
      {
      if (md->noteol) MRRETURN(MATCH_NOMATCH);
      SCHECK_PARTIAL();
      }
    ecode++;
    break;

    /* Not multiline mode: assert before a terminating newline or before end of
    subject unless noteol is set. */

    case OP_DOLL:
    if (md->noteol) MRRETURN(MATCH_NOMATCH);
    if (!md->endonly) goto ASSERT_NL_OR_EOS;

    /* ... else fall through for endonly */

    /* End of subject assertion (\z) */

    case OP_EOD:
    if (eptr < md->end_subject) MRRETURN(MATCH_NOMATCH);
    SCHECK_PARTIAL();
    ecode++;
    break;

    /* End of subject or ending \n assertion (\Z) */

    case OP_EODN:
    ASSERT_NL_OR_EOS:
    if (eptr < md->end_subject &&
        (!IS_NEWLINE(eptr) || eptr != md->end_subject - md->nllen))
      MRRETURN(MATCH_NOMATCH);

    /* Either at end of string or \n before end. */

    SCHECK_PARTIAL();
    ecode++;
    break;

    /* Word boundary assertions */

    case OP_NOT_WORD_BOUNDARY:
    case OP_WORD_BOUNDARY:
      {

      /* Find out if the previous and current characters are "word" characters.
      It takes a bit more work in UTF-8 mode. Characters > 255 are assumed to
      be "non-word" characters. Remember the earliest consulted character for
      partial matching. */

#ifdef SUPPORT_UTF8
      if (utf8)
        {
        /* Get status of previous character */

        if (eptr == md->start_subject) prev_is_word = FALSE; else
          {
          USPTR lastptr = eptr - 1;
          TBACKCHAR(lastptr);
          if (lastptr < md->start_used_ptr) md->start_used_ptr = lastptr;
          TGETCHAR(c, lastptr);
#ifdef SUPPORT_UCP
          if (md->use_ucp)
            {
            if (c == '_') prev_is_word = TRUE; else
              {
              int cat = UCD_CATEGORY(c);
              prev_is_word = (cat == ucp_L || cat == ucp_N);
              }
            }
          else
#endif
          prev_is_word = c < 256 && (md->ctypes[c] & ctype_word) != 0;
          }

        /* Get status of next character */

        if (eptr >= md->end_subject)
          {
          SCHECK_PARTIAL();
          cur_is_word = FALSE;
          }
        else
          {
          TGETCHAR(c, eptr);
#ifdef SUPPORT_UCP
          if (md->use_ucp)
            {
            if (c == '_') cur_is_word = TRUE; else
              {
              int cat = UCD_CATEGORY(c);
              cur_is_word = (cat == ucp_L || cat == ucp_N);
              }
            }
          else
#endif
          cur_is_word = c < 256 && (md->ctypes[c] & ctype_word) != 0;
          }
        }
      else
#endif

      /* Not in UTF-8 mode, but we may still have PCRE_UCP set, and for
      consistency with the behaviour of \w we do use it in this case. */

        {
        /* Get status of previous character */

        if (eptr == md->start_subject) prev_is_word = FALSE; else
          {
          if (eptr <= md->start_used_ptr) md->start_used_ptr = eptr - 1;
#ifdef SUPPORT_UCP
          if (md->use_ucp)
            {
            c = eptr[-1];
            if (c == '_') prev_is_word = TRUE; else
              {
              int cat = UCD_CATEGORY(c);
              prev_is_word = (cat == ucp_L || cat == ucp_N);
              }
            }
          else
#endif
          prev_is_word = ((md->ctypes[eptr[-1]] & ctype_word) != 0);
          }

        /* Get status of next character */

        if (eptr >= md->end_subject)
          {
          SCHECK_PARTIAL();
          cur_is_word = FALSE;
          }
        else
#ifdef SUPPORT_UCP
        if (md->use_ucp)
          {
          c = *eptr;
          if (c == '_') cur_is_word = TRUE; else
            {
            int cat = UCD_CATEGORY(c);
            cur_is_word = (cat == ucp_L || cat == ucp_N);
            }
          }
        else
#endif
        cur_is_word = ((md->ctypes[*eptr] & ctype_word) != 0);
        }

      /* Now see if the situation is what we want */

      if ((*ecode++ == OP_WORD_BOUNDARY)?
           cur_is_word == prev_is_word : cur_is_word != prev_is_word)
        MRRETURN(MATCH_NOMATCH);
      }
    break;

    /* Match a single character type; inline for speed */

    case OP_ANY:
    if (IS_NEWLINE(eptr)) MRRETURN(MATCH_NOMATCH);
    /* Fall through */

    case OP_ALLANY:
    if (eptr >= md->end_subject)   /* DO NOT merge the eptr++ here; it must */
      {                            /* not be updated before SCHECK_PARTIAL. */
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    eptr++;
#ifdef SUPPORT_UTF8
    if (utf8) TSKIPCHAR(eptr);
#endif
    ecode++;
    break;

    /* Match a single byte, even in UTF-8 mode. This opcode really does match
    any byte, even newline, independent of the setting of PCRE_DOTALL. */

    case OP_ANYBYTE:
    if (eptr >= md->end_subject)   /* DO NOT merge the eptr++ here; it must */
      {                            /* not be updated before SCHECK_PARTIAL. */
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    eptr++;
    ecode++;
    break;

    case OP_NOT_DIGIT:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (
#ifdef SUPPORT_UTF8
       c < 256 &&
#endif
       (md->ctypes[c] & ctype_digit) != 0
       )
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    case OP_DIGIT:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (
#ifdef SUPPORT_UTF8
       c >= 256 ||
#endif
       (md->ctypes[c] & ctype_digit) == 0
       )
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    case OP_NOT_WHITESPACE:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (
#ifdef SUPPORT_UTF8
       c < 256 &&
#endif
       (md->ctypes[c] & ctype_space) != 0
       )
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    case OP_WHITESPACE:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (
#ifdef SUPPORT_UTF8
       c >= 256 ||
#endif
       (md->ctypes[c] & ctype_space) == 0
       )
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    case OP_NOT_WORDCHAR:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (
#ifdef SUPPORT_UTF8
       c < 256 &&
#endif
       (md->ctypes[c] & ctype_word) != 0
       )
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    case OP_WORDCHAR:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (
#ifdef SUPPORT_UTF8
       c >= 256 ||
#endif
       (md->ctypes[c] & ctype_word) == 0
       )
      MRRETURN(MATCH_NOMATCH);
    ecode++;
    break;

    case OP_ANYNL:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    switch(c)
      {
      default: MRRETURN(MATCH_NOMATCH);

      case 0x000d:
      if (eptr < md->end_subject && *eptr == 0x0a) eptr++;
      break;

      case 0x000a:
      break;

      case 0x000b:
      case 0x000c:
      case 0x0085:
      case 0x2028:
      case 0x2029:
      if (md->bsr_anycrlf) MRRETURN(MATCH_NOMATCH);
      break;
      }
    ecode++;
    break;

    case OP_NOT_HSPACE:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    switch(c)
      {
      default: break;
      case 0x09:      /* HT */
      case 0x20:      /* SPACE */
      case 0xa0:      /* NBSP */
      case 0x1680:    /* OGHAM SPACE MARK */
      case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
      case 0x2000:    /* EN QUAD */
      case 0x2001:    /* EM QUAD */
      case 0x2002:    /* EN SPACE */
      case 0x2003:    /* EM SPACE */
      case 0x2004:    /* THREE-PER-EM SPACE */
      case 0x2005:    /* FOUR-PER-EM SPACE */
      case 0x2006:    /* SIX-PER-EM SPACE */
      case 0x2007:    /* FIGURE SPACE */
      case 0x2008:    /* PUNCTUATION SPACE */
      case 0x2009:    /* THIN SPACE */
      case 0x200A:    /* HAIR SPACE */
      case 0x202f:    /* NARROW NO-BREAK SPACE */
      case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
      case 0x3000:    /* IDEOGRAPHIC SPACE */
      MRRETURN(MATCH_NOMATCH);
      }
    ecode++;
    break;

    case OP_HSPACE:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    switch(c)
      {
      default: MRRETURN(MATCH_NOMATCH);
      case 0x09:      /* HT */
      case 0x20:      /* SPACE */
      case 0xa0:      /* NBSP */
      case 0x1680:    /* OGHAM SPACE MARK */
      case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
      case 0x2000:    /* EN QUAD */
      case 0x2001:    /* EM QUAD */
      case 0x2002:    /* EN SPACE */
      case 0x2003:    /* EM SPACE */
      case 0x2004:    /* THREE-PER-EM SPACE */
      case 0x2005:    /* FOUR-PER-EM SPACE */
      case 0x2006:    /* SIX-PER-EM SPACE */
      case 0x2007:    /* FIGURE SPACE */
      case 0x2008:    /* PUNCTUATION SPACE */
      case 0x2009:    /* THIN SPACE */
      case 0x200A:    /* HAIR SPACE */
      case 0x202f:    /* NARROW NO-BREAK SPACE */
      case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
      case 0x3000:    /* IDEOGRAPHIC SPACE */
      break;
      }
    ecode++;
    break;

    case OP_NOT_VSPACE:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    switch(c)
      {
      default: break;
      case 0x0a:      /* LF */
      case 0x0b:      /* VT */
      case 0x0c:      /* FF */
      case 0x0d:      /* CR */
      case 0x85:      /* NEL */
      case 0x2028:    /* LINE SEPARATOR */
      case 0x2029:    /* PARAGRAPH SEPARATOR */
      MRRETURN(MATCH_NOMATCH);
      }
    ecode++;
    break;

    case OP_VSPACE:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    switch(c)
      {
      default: MRRETURN(MATCH_NOMATCH);
      case 0x0a:      /* LF */
      case 0x0b:      /* VT */
      case 0x0c:      /* FF */
      case 0x0d:      /* CR */
      case 0x85:      /* NEL */
      case 0x2028:    /* LINE SEPARATOR */
      case 0x2029:    /* PARAGRAPH SEPARATOR */
      break;
      }
    ecode++;
    break;

#ifdef SUPPORT_UCP
    /* Check the next character by Unicode property. We will get here only
    if the support is in the binary; otherwise a compile-time error occurs. */

    case OP_PROP:
    case OP_NOTPROP:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
      {
      const ucd_record *prop = GET_UCD(c);

      switch(ecode[1])
        {
        case PT_ANY:
        if (op == OP_NOTPROP) MRRETURN(MATCH_NOMATCH);
        break;

        case PT_LAMP:
        if ((prop->chartype == ucp_Lu ||
             prop->chartype == ucp_Ll ||
             prop->chartype == ucp_Lt) == (op == OP_NOTPROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        case PT_GC:
        if ((ecode[2] != _pcre_ucp_gentype[prop->chartype]) == (op == OP_PROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        case PT_PC:
        if ((ecode[2] != prop->chartype) == (op == OP_PROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        case PT_SC:
        if ((ecode[2] != prop->script) == (op == OP_PROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        /* These are specials */

        case PT_ALNUM:
        if ((_pcre_ucp_gentype[prop->chartype] == ucp_L ||
             _pcre_ucp_gentype[prop->chartype] == ucp_N) == (op == OP_NOTPROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        case PT_SPACE:    /* Perl space */
        if ((_pcre_ucp_gentype[prop->chartype] == ucp_Z ||
             c == CHAR_HT || c == CHAR_NL || c == CHAR_FF || c == CHAR_CR)
               == (op == OP_NOTPROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        case PT_PXSPACE:  /* POSIX space */
        if ((_pcre_ucp_gentype[prop->chartype] == ucp_Z ||
             c == CHAR_HT || c == CHAR_NL || c == CHAR_VT ||
             c == CHAR_FF || c == CHAR_CR)
               == (op == OP_NOTPROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        case PT_WORD:
        if ((_pcre_ucp_gentype[prop->chartype] == ucp_L ||
             _pcre_ucp_gentype[prop->chartype] == ucp_N ||
             c == CHAR_UNDERSCORE) == (op == OP_NOTPROP))
          MRRETURN(MATCH_NOMATCH);
        break;

        /* This should never occur */

        default:
        RRETURN(PCRE_ERROR_INTERNAL);
        }

      ecode += 3;
      }
    break;

    /* Match an extended Unicode sequence. We will get here only if the support
    is in the binary; otherwise a compile-time error occurs. */

    case OP_EXTUNI:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    TGETCHARINCTEST(c, eptr);
    if (UCD_CATEGORY(c) == ucp_M) MRRETURN(MATCH_NOMATCH);
    while (eptr < md->end_subject)
      {
      int len = 1;
      if (!utf8) c = *eptr; else { TGETCHARLEN(c, eptr, len); }
      if (UCD_CATEGORY(c) != ucp_M) break;
      eptr += len;
      }
    ecode++;
    break;
#endif


    /* Match a back reference, possibly repeatedly. Look past the end of the
    item to see if there is repeat information following. The code is similar
    to that for character classes, but repeated for efficiency. Then obey
    similar code to character type repeats - written out again for speed.
    However, if the referenced string is the empty string, always treat
    it as matched, any number of times (otherwise there could be infinite
    loops). */

    case OP_REF:
    case OP_REFI:
    caseless = op == OP_REFI;
    offset = GET2(ecode, 1) << 1;               /* Doubled ref number */
    ecode += 3;

    /* If the reference is unset, there are two possibilities:

    (a) In the default, Perl-compatible state, set the length negative;
    this ensures that every attempt at a match fails. We can't just fail
    here, because of the possibility of quantifiers with zero minima.

    (b) If the JavaScript compatibility flag is set, set the length to zero
    so that the back reference matches an empty string.

    Otherwise, set the length to the length of what was matched by the
    referenced subpattern. */

    if (offset >= offset_top || md->offset_vector[offset] < 0)
      length = (md->jscript_compat)? 0 : -1;
    else
      length = md->offset_vector[offset+1] - md->offset_vector[offset];

    /* Set up for repetition, or handle the non-repeated case */

    switch (*ecode)
      {
      case OP_CRSTAR:
      case OP_CRMINSTAR:
      case OP_CRPLUS:
      case OP_CRMINPLUS:
      case OP_CRQUERY:
      case OP_CRMINQUERY:
      c = *ecode++ - OP_CRSTAR;
      minimize = (c & 1) != 0;
      min = rep_min[c];                 /* Pick up values from tables; */
      max = rep_max[c];                 /* zero for max => infinity */
      if (max == 0) max = INT_MAX;
      break;

      case OP_CRRANGE:
      case OP_CRMINRANGE:
      minimize = (*ecode == OP_CRMINRANGE);
      min = GET2(ecode, 1);
      max = GET2(ecode, 3);
      if (max == 0) max = INT_MAX;
      ecode += 5;
      break;

      default:               /* No repeat follows */
      if ((length = match_ref(offset, eptr, length, md, caseless)) < 0)
        {
        CHECK_PARTIAL();
        MRRETURN(MATCH_NOMATCH);
        }
      eptr += length;
      continue;              /* With the main loop */
      }

    /* Handle repeated back references. If the length of the reference is
    zero, just continue with the main loop. */

    if (length == 0) continue;

    /* First, ensure the minimum number of matches are present. We get back
    the length of the reference string explicitly rather than passing the
    address of eptr, so that eptr can be a register variable. */

    for (i = 1; i <= min; i++)
      {
      int slength;
      if ((slength = match_ref(offset, eptr, length, md, caseless)) < 0)
        {
        CHECK_PARTIAL();
        MRRETURN(MATCH_NOMATCH);
        }
      eptr += slength;
      }

    /* If min = max, continue at the same level without recursion.
    They are not both allowed to be zero. */

    if (min == max) continue;

    /* If minimizing, keep trying and advancing the pointer */

    if (minimize)
      {
      for (fi = min;; fi++)
        {
        int slength;
        RMATCH(eptr, ecode, offset_top, md, eptrb, RM14);
        if (rrc != MATCH_NOMATCH) RRETURN(rrc);
        if (fi >= max) MRRETURN(MATCH_NOMATCH);
        if ((slength = match_ref(offset, eptr, length, md, caseless)) < 0)
          {
          CHECK_PARTIAL();
          MRRETURN(MATCH_NOMATCH);
          }
        eptr += slength;
        }
      /* Control never gets here */
      }

    /* If maximizing, find the longest string and work backwards */

    else
      {
      pp = eptr;
      for (i = min; i < max; i++)
        {
        int slength;
        if ((slength = match_ref(offset, eptr, length, md, caseless)) < 0)
          {
          CHECK_PARTIAL();
          break;
          }
        eptr += slength;
        }
      while (eptr >= pp)
        {
        RMATCH(eptr, ecode, offset_top, md, eptrb, RM15);
        if (rrc != MATCH_NOMATCH) RRETURN(rrc);
        eptr -= length;
        }
      MRRETURN(MATCH_NOMATCH);
      }
    /* Control never gets here */

    /* Match a bit-mapped character class, possibly repeatedly. This op code is
    used when all the characters in the class have values in the range 0-255,
    and either the matching is caseful, or the characters are in the range
    0-127 when UTF-8 processing is enabled. The only difference between
    OP_CLASS and OP_NCLASS occurs when a data character outside the range is
    encountered.

    First, look past the end of the item to see if there is repeat information
    following. Then obey similar code to character type repeats - written out
    again for speed. */

    case OP_NCLASS:
    case OP_CLASS:
      {
      data = ecode + 1;                /* Save for matching */
      ecode += 33;                     /* Advance past the item */

      switch (*ecode)
        {
        case OP_CRSTAR:
        case OP_CRMINSTAR:
        case OP_CRPLUS:
        case OP_CRMINPLUS:
        case OP_CRQUERY:
        case OP_CRMINQUERY:
        c = *ecode++ - OP_CRSTAR;
        minimize = (c & 1) != 0;
        min = rep_min[c];                 /* Pick up values from tables; */
        max = rep_max[c];                 /* zero for max => infinity */
        if (max == 0) max = INT_MAX;
        break;

        case OP_CRRANGE:
        case OP_CRMINRANGE:
        minimize = (*ecode == OP_CRMINRANGE);
        min = GET2(ecode, 1);
        max = GET2(ecode, 3);
        if (max == 0) max = INT_MAX;
        ecode += 5;
        break;

        default:               /* No repeat follows */
        min = max = 1;
        break;
        }

      /* First, ensure the minimum number of matches are present. */

#ifdef SUPPORT_UTF8
      /* UTF-8 mode */
      if (utf8)
        {
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          if (c > 255)
            {
            if (op == OP_CLASS) MRRETURN(MATCH_NOMATCH);
            }
          else
            {
            if ((data[c/8] & (1 << (c&7))) == 0) MRRETURN(MATCH_NOMATCH);
            }
          }
        }
      else
#endif
      /* Not UTF-8 mode */
        {
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          c = *eptr++;
          if ((data[c/8] & (1 << (c&7))) == 0) MRRETURN(MATCH_NOMATCH);
          }
        }

      /* If max == min we can continue with the main loop without the
      need to recurse. */

      if (min == max) continue;

      /* If minimizing, keep testing the rest of the expression and advancing
      the pointer while it matches the class. */

      if (minimize)
        {
#ifdef SUPPORT_UTF8
        /* UTF-8 mode */
        if (utf8)
          {
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM16);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINC(c, eptr);
            if (c > 255)
              {
              if (op == OP_CLASS) MRRETURN(MATCH_NOMATCH);
              }
            else
              {
              if ((data[c/8] & (1 << (c&7))) == 0) MRRETURN(MATCH_NOMATCH);
              }
            }
          }
        else
#endif
        /* Not UTF-8 mode */
          {
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM17);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            c = *eptr++;
            if ((data[c/8] & (1 << (c&7))) == 0) MRRETURN(MATCH_NOMATCH);
            }
          }
        /* Control never gets here */
        }

      /* If maximizing, find the longest possible run, then work backwards. */

      else
        {
        pp = eptr;

#ifdef SUPPORT_UTF8
        /* UTF-8 mode */
        if (utf8)
          {
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c > 255)
              {
              if (op == OP_CLASS) break;
              }
            else
              {
              if ((data[c/8] & (1 << (c&7))) == 0) break;
              }
            eptr += len;
            }
          for (;;)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM18);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (eptr-- == pp) break;        /* Stop if tried at original pos */
            TBACKCHAR(eptr);
            }
          }
        else
#endif
          /* Not UTF-8 mode */
          {
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            c = *eptr;
            if ((data[c/8] & (1 << (c&7))) == 0) break;
            eptr++;
            }
          while (eptr >= pp)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM19);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            eptr--;
            }
          }

        MRRETURN(MATCH_NOMATCH);
        }
      }
    /* Control never gets here */


    /* Match an extended character class. This opcode is encountered only
    when UTF-8 mode mode is supported. Nevertheless, we may not be in UTF-8
    mode, because Unicode properties are supported in non-UTF-8 mode. */

#ifdef SUPPORT_UTF8
    case OP_XCLASS:
      {
      data = ecode + 1 + LINK_SIZE;                /* Save for matching */
      ecode += GET(ecode, 1);                      /* Advance past the item */

      switch (*ecode)
        {
        case OP_CRSTAR:
        case OP_CRMINSTAR:
        case OP_CRPLUS:
        case OP_CRMINPLUS:
        case OP_CRQUERY:
        case OP_CRMINQUERY:
        c = *ecode++ - OP_CRSTAR;
        minimize = (c & 1) != 0;
        min = rep_min[c];                 /* Pick up values from tables; */
        max = rep_max[c];                 /* zero for max => infinity */
        if (max == 0) max = INT_MAX;
        break;

        case OP_CRRANGE:
        case OP_CRMINRANGE:
        minimize = (*ecode == OP_CRMINRANGE);
        min = GET2(ecode, 1);
        max = GET2(ecode, 3);
        if (max == 0) max = INT_MAX;
        ecode += 5;
        break;

        default:               /* No repeat follows */
        min = max = 1;
        break;
        }

      /* First, ensure the minimum number of matches are present. */

      for (i = 1; i <= min; i++)
        {
        if (eptr >= md->end_subject)
          {
          SCHECK_PARTIAL();
          MRRETURN(MATCH_NOMATCH);
          }
        TGETCHARINCTEST(c, eptr);
        if (!_pcre_xclass(c, data)) MRRETURN(MATCH_NOMATCH);
        }

      /* If max == min we can continue with the main loop without the
      need to recurse. */

      if (min == max) continue;

      /* If minimizing, keep testing the rest of the expression and advancing
      the pointer while it matches the class. */

      if (minimize)
        {
        for (fi = min;; fi++)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM20);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (fi >= max) MRRETURN(MATCH_NOMATCH);
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINCTEST(c, eptr);
          if (!_pcre_xclass(c, data)) MRRETURN(MATCH_NOMATCH);
          }
        /* Control never gets here */
        }

      /* If maximizing, find the longest possible run, then work backwards. */

      else
        {
        pp = eptr;
        for (i = min; i < max; i++)
          {
          int len = 1;
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            break;
            }
          TGETCHARLENTEST(c, eptr, len);
          if (!_pcre_xclass(c, data)) break;
          eptr += len;
          }
        for(;;)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM21);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (eptr-- == pp) break;        /* Stop if tried at original pos */
          if (utf8) TBACKCHAR(eptr);
          }
        MRRETURN(MATCH_NOMATCH);
        }

      /* Control never gets here */
      }
#endif    /* End of XCLASS */

    /* Match a single character, casefully */

    case OP_CHAR:

    /* AutoHotkey: Partial matches with incomplete multi-byte chars aren't supported
    since ecode and eptr have different encodings. */

    if (md->end_subject - eptr < 1)
      {
      SCHECK_PARTIAL();            /* This one can use SCHECK_PARTIAL() */
      MRRETURN(MATCH_NOMATCH);
      }

    /* Supports both UTF-8 and non-UTF-8 modes: */

    else
      {
      unsigned int dc;

      ecode++;
      GETCHARINC(fc, ecode);
      TGETCHARINC(dc, eptr);

      if (fc != dc) MRRETURN(MATCH_NOMATCH);

      }

    break;

    /* Match a single character, caselessly */

    case OP_CHARI:

    if (md->end_subject - eptr < 1)
      {
      SCHECK_PARTIAL();            /* This one can use SCHECK_PARTIAL() */
      MRRETURN(MATCH_NOMATCH);
      }

#ifdef SUPPORT_UTF8
    if (utf8)
      {
      unsigned int dc;

      ecode++;
      GETCHARINC(fc, ecode);
      TGETCHARINC(dc, eptr);

      /* If the characters' values are < 128, we can use the fast lookup table. */

      if (fc < 128 && dc < 128)
        {
        if (md->lcc[fc] != md->lcc[dc]) MRRETURN(MATCH_NOMATCH);
        }

      else
        {

        /* If we have Unicode property support, we can use it to test the other
        case of the character, if there is one. */

        if (fc != dc)
          {
#ifdef SUPPORT_UCP
          if (dc != UCD_OTHERCASE(fc))
#endif
            MRRETURN(MATCH_NOMATCH);
          }
        }
      }
    else
#endif   /* SUPPORT_UTF8 */

    /* Non-UTF-8 mode */
      {
      if (md->lcc[ecode[1]] != md->lcc[*eptr++]) MRRETURN(MATCH_NOMATCH);
      ecode += 2;
      }
    break;

    /* Match a single character repeatedly. */

    case OP_EXACT:
    case OP_EXACTI:
    min = max = GET2(ecode, 1);
    ecode += 3;
    goto REPEATCHAR;

    case OP_POSUPTO:
    case OP_POSUPTOI:
    possessive = TRUE;
    /* Fall through */

    case OP_UPTO:
    case OP_UPTOI:
    case OP_MINUPTO:
    case OP_MINUPTOI:
    min = 0;
    max = GET2(ecode, 1);
    minimize = *ecode == OP_MINUPTO || *ecode == OP_MINUPTOI;
    ecode += 3;
    goto REPEATCHAR;

    case OP_POSSTAR:
    case OP_POSSTARI:
    possessive = TRUE;
    min = 0;
    max = INT_MAX;
    ecode++;
    goto REPEATCHAR;

    case OP_POSPLUS:
    case OP_POSPLUSI:
    possessive = TRUE;
    min = 1;
    max = INT_MAX;
    ecode++;
    goto REPEATCHAR;

    case OP_POSQUERY:
    case OP_POSQUERYI:
    possessive = TRUE;
    min = 0;
    max = 1;
    ecode++;
    goto REPEATCHAR;

    case OP_STAR:
    case OP_STARI:
    case OP_MINSTAR:
    case OP_MINSTARI:
    case OP_PLUS:
    case OP_PLUSI:
    case OP_MINPLUS:
    case OP_MINPLUSI:
    case OP_QUERY:
    case OP_QUERYI:
    case OP_MINQUERY:
    case OP_MINQUERYI:
    c = *ecode++ - ((op < OP_STARI)? OP_STAR : OP_STARI);
    minimize = (c & 1) != 0;
    min = rep_min[c];                 /* Pick up values from tables; */
    max = rep_max[c];                 /* zero for max => infinity */
    if (max == 0) max = INT_MAX;

    /* Common code for all repeated single-character matches. */

    REPEATCHAR:
#ifdef SUPPORT_UTF8
    if (utf8)
      {
#ifdef PCRE_USE_UTF16
      utchar mcbuffer[2];
      charptr = mcbuffer;
      GETCHARINC(fc, ecode);
      length = _pcre_ord2utf16(fc, mcbuffer);
#else
      charptr = ecode;
      length = 1;
      GETCHARLEN(fc, ecode, length);
      ecode += length;
#endif
      
      /* Handle multibyte character matching specially here. There is
      support for caseless matching if UCP support is present. */

      if (fc > 127)
        {
#ifdef SUPPORT_UCP
        unsigned int othercase;
        if (op >= OP_STARI &&     /* Caseless */
            (othercase = UCD_OTHERCASE(fc)) != fc)
#ifdef PCRE_USE_UTF16
          oclength = _pcre_ord2utf16(othercase, occhars);
#else
          oclength = _pcre_ord2utf8(othercase, occhars);
#endif
        else oclength = 0;
#endif  /* SUPPORT_UCP */

        for (i = 1; i <= min; i++)
          {
          if (eptr <= md->end_subject - length &&
            tmemcmp(eptr, charptr, length) == 0) eptr += length;
#ifdef SUPPORT_UCP
          else if (oclength > 0 &&
                   eptr <= md->end_subject - oclength &&
                   tmemcmp(eptr, occhars, oclength) == 0) eptr += oclength;
#endif  /* SUPPORT_UCP */
          else
            {
            CHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          }

        if (min == max) continue;

        if (minimize)
          {
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM22);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr <= md->end_subject - length &&
              tmemcmp(eptr, charptr, length) == 0) eptr += length;
#ifdef SUPPORT_UCP
            else if (oclength > 0 &&
                     eptr <= md->end_subject - oclength &&
                     tmemcmp(eptr, occhars, oclength) == 0) eptr += oclength;
#endif  /* SUPPORT_UCP */
            else
              {
              CHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            }
          /* Control never gets here */
          }

        else  /* Maximize */
          {
          pp = eptr;
          for (i = min; i < max; i++)
            {
            if (eptr <= md->end_subject - length &&
                tmemcmp(eptr, charptr, length) == 0) eptr += length;
#ifdef SUPPORT_UCP
            else if (oclength > 0 &&
                     eptr <= md->end_subject - oclength &&
                     tmemcmp(eptr, occhars, oclength) == 0) eptr += oclength;
#endif  /* SUPPORT_UCP */
            else
              {
              CHECK_PARTIAL();
              break;
              }
            }

          if (possessive) continue;

          for(;;)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM23);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (eptr == pp) { MRRETURN(MATCH_NOMATCH); }
#ifdef SUPPORT_UCP
            eptr--;
            TBACKCHAR(eptr);
#else   /* without SUPPORT_UCP */
            eptr -= length;
#endif  /* SUPPORT_UCP */
            }
          }
        /* Control never gets here */
        }

      /* If the length of a UTF-8 character is 1, we fall through here, and
      obey the code as for non-UTF-8 characters below, though in this case the
      value of fc will always be < 128. */
      }
    else
#endif  /* SUPPORT_UTF8 */

    /* When not in UTF-8 mode, load a single-byte character. */

    fc = *ecode++;

    /* The value of fc at this point is always less than 256, though we may or
    may not be in UTF-8 mode. The code is duplicated for the caseless and
    caseful cases, for speed, since matching characters is likely to be quite
    common. First, ensure the minimum number of matches are present. If min =
    max, continue at the same level without recursing. Otherwise, if
    minimizing, keep trying the rest of the expression and advancing one
    matching character if failing, up to the maximum. Alternatively, if
    maximizing, find the maximum number of characters and work backwards. */

    DPRINTF(("matching %c{%d,%d} against subject %.*s\n", fc, min, max,
      max, eptr));

    if (op >= OP_STARI)  /* Caseless */
      {
      unsigned int dc;
      fc = md->lcc[fc];
      for (i = 1; i <= min; i++)
        {
        if (eptr >= md->end_subject)
          {
          SCHECK_PARTIAL();
          MRRETURN(MATCH_NOMATCH);
          }
        TGETCHARINC(dc, eptr);
        if (dc > 255 || fc != md->lcc[dc]) MRRETURN(MATCH_NOMATCH);
        }
      if (min == max) continue;
      if (minimize)
        {
        for (fi = min;; fi++)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM24);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (fi >= max) MRRETURN(MATCH_NOMATCH);
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(dc, eptr);
          if (dc > 255 || fc != md->lcc[dc]) MRRETURN(MATCH_NOMATCH);
          }
        /* Control never gets here */
        }
      else  /* Maximize */
        {
        pp = eptr;
        for (i = min; i < max; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            break;
            }
          length = 1;
          TGETCHARLEN(dc, eptr, length);
          if (dc > 255 || fc != md->lcc[dc]) break;
          eptr += length;
          }

        if (possessive) continue;

        while (eptr >= pp)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM25);
          eptr--;
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          }
        MRRETURN(MATCH_NOMATCH);
        }
      /* Control never gets here */
      }

    /* Caseful comparisons (includes all multi-byte characters) */

    else
      {
      for (i = 1; i <= min; i++)
        {
        if (eptr >= md->end_subject)
          {
          SCHECK_PARTIAL();
          MRRETURN(MATCH_NOMATCH);
          }
        if (fc != *eptr++) MRRETURN(MATCH_NOMATCH);
        }

      if (min == max) continue;

      if (minimize)
        {
        for (fi = min;; fi++)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM26);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (fi >= max) MRRETURN(MATCH_NOMATCH);
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (fc != *eptr++) MRRETURN(MATCH_NOMATCH);
          }
        /* Control never gets here */
        }
      else  /* Maximize */
        {
        pp = eptr;
        for (i = min; i < max; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            break;
            }
          if (fc != *eptr) break;
          eptr++;
          }
        if (possessive) continue;

        while (eptr >= pp)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM27);
          eptr--;
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          }
        MRRETURN(MATCH_NOMATCH);
        }
      }
    /* Control never gets here */

    /* Match a negated single one-byte character. The character we are
    checking can be multibyte. */

    case OP_NOT:
    case OP_NOTI:
    if (eptr >= md->end_subject)
      {
      SCHECK_PARTIAL();
      MRRETURN(MATCH_NOMATCH);
      }
    ecode++;
    TGETCHARINCTEST(c, eptr);
    if (op == OP_NOTI)         /* The caseless case */
      {
#ifdef SUPPORT_UTF8
      if (c < 256)
#endif
      c = md->lcc[c];
      if (md->lcc[*ecode++] == c) MRRETURN(MATCH_NOMATCH);
      }
    else    /* Caseful */
      {
      if (*ecode++ == c) MRRETURN(MATCH_NOMATCH);
      }
    break;

    /* Match a negated single one-byte character repeatedly. This is almost a
    repeat of the code for a repeated single character, but I haven't found a
    nice way of commoning these up that doesn't require a test of the
    positive/negative option for each character match. Maybe that wouldn't add
    very much to the time taken, but character matching *is* what this is all
    about... */

    case OP_NOTEXACT:
    case OP_NOTEXACTI:
    min = max = GET2(ecode, 1);
    ecode += 3;
    goto REPEATNOTCHAR;

    case OP_NOTUPTO:
    case OP_NOTUPTOI:
    case OP_NOTMINUPTO:
    case OP_NOTMINUPTOI:
    min = 0;
    max = GET2(ecode, 1);
    minimize = *ecode == OP_NOTMINUPTO || *ecode == OP_NOTMINUPTOI;
    ecode += 3;
    goto REPEATNOTCHAR;

    case OP_NOTPOSSTAR:
    case OP_NOTPOSSTARI:
    possessive = TRUE;
    min = 0;
    max = INT_MAX;
    ecode++;
    goto REPEATNOTCHAR;

    case OP_NOTPOSPLUS:
    case OP_NOTPOSPLUSI:
    possessive = TRUE;
    min = 1;
    max = INT_MAX;
    ecode++;
    goto REPEATNOTCHAR;

    case OP_NOTPOSQUERY:
    case OP_NOTPOSQUERYI:
    possessive = TRUE;
    min = 0;
    max = 1;
    ecode++;
    goto REPEATNOTCHAR;

    case OP_NOTPOSUPTO:
    case OP_NOTPOSUPTOI:
    possessive = TRUE;
    min = 0;
    max = GET2(ecode, 1);
    ecode += 3;
    goto REPEATNOTCHAR;

    case OP_NOTSTAR:
    case OP_NOTSTARI:
    case OP_NOTMINSTAR:
    case OP_NOTMINSTARI:
    case OP_NOTPLUS:
    case OP_NOTPLUSI:
    case OP_NOTMINPLUS:
    case OP_NOTMINPLUSI:
    case OP_NOTQUERY:
    case OP_NOTQUERYI:
    case OP_NOTMINQUERY:
    case OP_NOTMINQUERYI:
    c = *ecode++ - ((op >= OP_NOTSTARI)? OP_NOTSTARI: OP_NOTSTAR);
    minimize = (c & 1) != 0;
    min = rep_min[c];                 /* Pick up values from tables; */
    max = rep_max[c];                 /* zero for max => infinity */
    if (max == 0) max = INT_MAX;

    /* Common code for all repeated single-byte matches. */

    REPEATNOTCHAR:
    fc = *ecode++;

    /* The code is duplicated for the caseless and caseful cases, for speed,
    since matching characters is likely to be quite common. First, ensure the
    minimum number of matches are present. If min = max, continue at the same
    level without recursing. Otherwise, if minimizing, keep trying the rest of
    the expression and advancing one matching character if failing, up to the
    maximum. Alternatively, if maximizing, find the maximum number of
    characters and work backwards. */

    DPRINTF(("negative matching %c{%d,%d} against subject %.*s\n", fc, min, max,
      max, eptr));

    if (op >= OP_NOTSTARI)     /* Caseless */
      {
      fc = md->lcc[fc];

#ifdef SUPPORT_UTF8
      /* UTF-8 mode */
      if (utf8)
        {
        register unsigned int d;
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(d, eptr);
          if (d < 256) d = md->lcc[d];
          if (fc == d) MRRETURN(MATCH_NOMATCH);
          }
        }
      else
#endif

      /* Not UTF-8 mode */
        {
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (fc == md->lcc[*eptr++]) MRRETURN(MATCH_NOMATCH);
          }
        }

      if (min == max) continue;

      if (minimize)
        {
#ifdef SUPPORT_UTF8
        /* UTF-8 mode */
        if (utf8)
          {
          register unsigned int d;
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM28);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINC(d, eptr);
            if (d < 256) d = md->lcc[d];
            if (fc == d) MRRETURN(MATCH_NOMATCH);
            }
          }
        else
#endif
        /* Not UTF-8 mode */
          {
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM29);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            if (fc == md->lcc[*eptr++]) MRRETURN(MATCH_NOMATCH);
            }
          }
        /* Control never gets here */
        }

      /* Maximize case */

      else
        {
        pp = eptr;

#ifdef SUPPORT_UTF8
        /* UTF-8 mode */
        if (utf8)
          {
          register unsigned int d;
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(d, eptr, len);
            if (d < 256) d = md->lcc[d];
            if (fc == d) break;
            eptr += len;
            }
        if (possessive) continue;
        for(;;)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM30);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (eptr-- == pp) break;        /* Stop if tried at original pos */
            TBACKCHAR(eptr);
            }
          }
        else
#endif
        /* Not UTF-8 mode */
          {
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if (fc == md->lcc[*eptr]) break;
            eptr++;
            }
          if (possessive) continue;
          while (eptr >= pp)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM31);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            eptr--;
            }
          }

        MRRETURN(MATCH_NOMATCH);
        }
      /* Control never gets here */
      }

    /* Caseful comparisons */

    else
      {
#ifdef SUPPORT_UTF8
      /* UTF-8 mode */
      if (utf8)
        {
        register unsigned int d;
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(d, eptr);
          if (fc == d) MRRETURN(MATCH_NOMATCH);
          }
        }
      else
#endif
      /* Not UTF-8 mode */
        {
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (fc == *eptr++) MRRETURN(MATCH_NOMATCH);
          }
        }

      if (min == max) continue;

      if (minimize)
        {
#ifdef SUPPORT_UTF8
        /* UTF-8 mode */
        if (utf8)
          {
          register unsigned int d;
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM32);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINC(d, eptr);
            if (fc == d) MRRETURN(MATCH_NOMATCH);
            }
          }
        else
#endif
        /* Not UTF-8 mode */
          {
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM33);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            if (fc == *eptr++) MRRETURN(MATCH_NOMATCH);
            }
          }
        /* Control never gets here */
        }

      /* Maximize case */

      else
        {
        pp = eptr;

#ifdef SUPPORT_UTF8
        /* UTF-8 mode */
        if (utf8)
          {
          register unsigned int d;
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(d, eptr, len);
            if (fc == d) break;
            eptr += len;
            }
          if (possessive) continue;
          for(;;)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM34);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (eptr-- == pp) break;        /* Stop if tried at original pos */
            TBACKCHAR(eptr);
            }
          }
        else
#endif
        /* Not UTF-8 mode */
          {
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if (fc == *eptr) break;
            eptr++;
            }
          if (possessive) continue;
          while (eptr >= pp)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM35);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            eptr--;
            }
          }

        MRRETURN(MATCH_NOMATCH);
        }
      }
    /* Control never gets here */

    /* Match a single character type repeatedly; several different opcodes
    share code. This is very similar to the code for single characters, but we
    repeat it in the interests of efficiency. */

    case OP_TYPEEXACT:
    min = max = GET2(ecode, 1);
    minimize = TRUE;
    ecode += 3;
    goto REPEATTYPE;

    case OP_TYPEUPTO:
    case OP_TYPEMINUPTO:
    min = 0;
    max = GET2(ecode, 1);
    minimize = *ecode == OP_TYPEMINUPTO;
    ecode += 3;
    goto REPEATTYPE;

    case OP_TYPEPOSSTAR:
    possessive = TRUE;
    min = 0;
    max = INT_MAX;
    ecode++;
    goto REPEATTYPE;

    case OP_TYPEPOSPLUS:
    possessive = TRUE;
    min = 1;
    max = INT_MAX;
    ecode++;
    goto REPEATTYPE;

    case OP_TYPEPOSQUERY:
    possessive = TRUE;
    min = 0;
    max = 1;
    ecode++;
    goto REPEATTYPE;

    case OP_TYPEPOSUPTO:
    possessive = TRUE;
    min = 0;
    max = GET2(ecode, 1);
    ecode += 3;
    goto REPEATTYPE;

    case OP_TYPESTAR:
    case OP_TYPEMINSTAR:
    case OP_TYPEPLUS:
    case OP_TYPEMINPLUS:
    case OP_TYPEQUERY:
    case OP_TYPEMINQUERY:
    c = *ecode++ - OP_TYPESTAR;
    minimize = (c & 1) != 0;
    min = rep_min[c];                 /* Pick up values from tables; */
    max = rep_max[c];                 /* zero for max => infinity */
    if (max == 0) max = INT_MAX;

    /* Common code for all repeated single character type matches. Note that
    in UTF-8 mode, '.' matches a character of any length, but for the other
    character types, the valid characters are all one-byte long. */

    REPEATTYPE:
    ctype = *ecode++;      /* Code for the character type */

#ifdef SUPPORT_UCP
    if (ctype == OP_PROP || ctype == OP_NOTPROP)
      {
      prop_fail_result = ctype == OP_NOTPROP;
      prop_type = *ecode++;
      prop_value = *ecode++;
      }
    else prop_type = -1;
#endif

    /* First, ensure the minimum number of matches are present. Use inline
    code for maximizing the speed, and do the type test once at the start
    (i.e. keep it out of the loop). Separate the UTF-8 code completely as that
    is tidier. Also separate the UCP code, which can be the same for both UTF-8
    and single-bytes. */

    if (min > 0)
      {
#ifdef SUPPORT_UCP
      if (prop_type >= 0)
        {
        switch(prop_type)
          {
          case PT_ANY:
          if (prop_fail_result) MRRETURN(MATCH_NOMATCH);
          for (i = 1; i <= min; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            }
          break;

          case PT_LAMP:
          for (i = 1; i <= min; i++)
            {
            int chartype;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            chartype = UCD_CHARTYPE(c);
            if ((chartype == ucp_Lu ||
                 chartype == ucp_Ll ||
                 chartype == ucp_Lt) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_GC:
          for (i = 1; i <= min; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CATEGORY(c) == prop_value) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_PC:
          for (i = 1; i <= min; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CHARTYPE(c) == prop_value) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_SC:
          for (i = 1; i <= min; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_SCRIPT(c) == prop_value) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_ALNUM:
          for (i = 1; i <= min; i++)
            {
            int category;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            category = UCD_CATEGORY(c);
            if ((category == ucp_L || category == ucp_N) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_SPACE:    /* Perl space */
          for (i = 1; i <= min; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CATEGORY(c) == ucp_Z || c == CHAR_HT || c == CHAR_NL ||
                 c == CHAR_FF || c == CHAR_CR)
                   == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_PXSPACE:  /* POSIX space */
          for (i = 1; i <= min; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CATEGORY(c) == ucp_Z || c == CHAR_HT || c == CHAR_NL ||
                 c == CHAR_VT || c == CHAR_FF || c == CHAR_CR)
                   == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          case PT_WORD:
          for (i = 1; i <= min; i++)
            {
            int category;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            category = UCD_CATEGORY(c);
            if ((category == ucp_L || category == ucp_N || c == CHAR_UNDERSCORE)
                   == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          break;

          /* This should not occur */

          default:
          RRETURN(PCRE_ERROR_INTERNAL);
          }
        }

      /* Match extended Unicode sequences. We will get here only if the
      support is in the binary; otherwise a compile-time error occurs. */

      else if (ctype == OP_EXTUNI)
        {
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINCTEST(c, eptr);
          if (UCD_CATEGORY(c) == ucp_M) MRRETURN(MATCH_NOMATCH);
          while (eptr < md->end_subject)
            {
            int len = 1;
            if (!utf8) c = *eptr; else { TGETCHARLEN(c, eptr, len); }
            if (UCD_CATEGORY(c) != ucp_M) break;
            eptr += len;
            }
          }
        }

      else
#endif     /* SUPPORT_UCP */

/* Handle all other cases when the coding is UTF-8 */

#ifdef SUPPORT_UTF8
      if (utf8) switch(ctype)
        {
        case OP_ANY:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (IS_NEWLINE(eptr)) MRRETURN(MATCH_NOMATCH);
          eptr++;
          TSKIPCHAR(eptr);
          }
        break;

        case OP_ALLANY:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          eptr++;
          TSKIPCHAR(eptr);
          }
        break;

        case OP_ANYBYTE:
        if (eptr > md->end_subject - min) MRRETURN(MATCH_NOMATCH);
        eptr += min;
        break;

        case OP_ANYNL:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          switch(c)
            {
            default: MRRETURN(MATCH_NOMATCH);

            case 0x000d:
            if (eptr < md->end_subject && *eptr == 0x0a) eptr++;
            break;

            case 0x000a:
            break;

            case 0x000b:
            case 0x000c:
            case 0x0085:
            case 0x2028:
            case 0x2029:
            if (md->bsr_anycrlf) MRRETURN(MATCH_NOMATCH);
            break;
            }
          }
        break;

        case OP_NOT_HSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          switch(c)
            {
            default: break;
            case 0x09:      /* HT */
            case 0x20:      /* SPACE */
            case 0xa0:      /* NBSP */
            case 0x1680:    /* OGHAM SPACE MARK */
            case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
            case 0x2000:    /* EN QUAD */
            case 0x2001:    /* EM QUAD */
            case 0x2002:    /* EN SPACE */
            case 0x2003:    /* EM SPACE */
            case 0x2004:    /* THREE-PER-EM SPACE */
            case 0x2005:    /* FOUR-PER-EM SPACE */
            case 0x2006:    /* SIX-PER-EM SPACE */
            case 0x2007:    /* FIGURE SPACE */
            case 0x2008:    /* PUNCTUATION SPACE */
            case 0x2009:    /* THIN SPACE */
            case 0x200A:    /* HAIR SPACE */
            case 0x202f:    /* NARROW NO-BREAK SPACE */
            case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
            case 0x3000:    /* IDEOGRAPHIC SPACE */
            MRRETURN(MATCH_NOMATCH);
            }
          }
        break;

        case OP_HSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          switch(c)
            {
            default: MRRETURN(MATCH_NOMATCH);
            case 0x09:      /* HT */
            case 0x20:      /* SPACE */
            case 0xa0:      /* NBSP */
            case 0x1680:    /* OGHAM SPACE MARK */
            case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
            case 0x2000:    /* EN QUAD */
            case 0x2001:    /* EM QUAD */
            case 0x2002:    /* EN SPACE */
            case 0x2003:    /* EM SPACE */
            case 0x2004:    /* THREE-PER-EM SPACE */
            case 0x2005:    /* FOUR-PER-EM SPACE */
            case 0x2006:    /* SIX-PER-EM SPACE */
            case 0x2007:    /* FIGURE SPACE */
            case 0x2008:    /* PUNCTUATION SPACE */
            case 0x2009:    /* THIN SPACE */
            case 0x200A:    /* HAIR SPACE */
            case 0x202f:    /* NARROW NO-BREAK SPACE */
            case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
            case 0x3000:    /* IDEOGRAPHIC SPACE */
            break;
            }
          }
        break;

        case OP_NOT_VSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          switch(c)
            {
            default: break;
            case 0x0a:      /* LF */
            case 0x0b:      /* VT */
            case 0x0c:      /* FF */
            case 0x0d:      /* CR */
            case 0x85:      /* NEL */
            case 0x2028:    /* LINE SEPARATOR */
            case 0x2029:    /* PARAGRAPH SEPARATOR */
            MRRETURN(MATCH_NOMATCH);
            }
          }
        break;

        case OP_VSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          switch(c)
            {
            default: MRRETURN(MATCH_NOMATCH);
            case 0x0a:      /* LF */
            case 0x0b:      /* VT */
            case 0x0c:      /* FF */
            case 0x0d:      /* CR */
            case 0x85:      /* NEL */
            case 0x2028:    /* LINE SEPARATOR */
            case 0x2029:    /* PARAGRAPH SEPARATOR */
            break;
            }
          }
        break;

        case OP_NOT_DIGIT:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINC(c, eptr);
          if (c < 128 && (md->ctypes[c] & ctype_digit) != 0)
            MRRETURN(MATCH_NOMATCH);
          }
        break;

        case OP_DIGIT:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (*eptr >= 128 || (md->ctypes[*eptr++] & ctype_digit) == 0)
            MRRETURN(MATCH_NOMATCH);
          /* No need to skip more bytes - we know it's a 1-byte character */
          }
        break;

        case OP_NOT_WHITESPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (*eptr < 128 && (md->ctypes[*eptr] & ctype_space) != 0)
            MRRETURN(MATCH_NOMATCH);
          eptr++;
          TSKIPCHAR(eptr);
          }
        break;

        case OP_WHITESPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (*eptr >= 128 || (md->ctypes[*eptr++] & ctype_space) == 0)
            MRRETURN(MATCH_NOMATCH);
          /* No need to skip more bytes - we know it's a 1-byte character */
          }
        break;

        case OP_NOT_WORDCHAR:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (*eptr < 128 && (md->ctypes[*eptr] & ctype_word) != 0)
            MRRETURN(MATCH_NOMATCH);
          eptr++;
          TSKIPCHAR(eptr);
          }
        break;

        case OP_WORDCHAR:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (*eptr >= 128 || (md->ctypes[*eptr++] & ctype_word) == 0)
            MRRETURN(MATCH_NOMATCH);
          /* No need to skip more bytes - we know it's a 1-byte character */
          }
        break;

        default:
        RRETURN(PCRE_ERROR_INTERNAL);
        }  /* End switch(ctype) */

      else
#endif     /* SUPPORT_UTF8 */

      /* Code for the non-UTF-8 case for minimum matching of operators other
      than OP_PROP and OP_NOTPROP. */

      switch(ctype)
        {
        case OP_ANY:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (IS_NEWLINE(eptr)) MRRETURN(MATCH_NOMATCH);
          eptr++;
          }
        break;

        case OP_ALLANY:
        if (eptr > md->end_subject - min)
          {
          SCHECK_PARTIAL();
          MRRETURN(MATCH_NOMATCH);
          }
        eptr += min;
        break;

        case OP_ANYBYTE:
        if (eptr > md->end_subject - min)
          {
          SCHECK_PARTIAL();
          MRRETURN(MATCH_NOMATCH);
          }
        eptr += min;
        break;

        case OP_ANYNL:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          switch(*eptr++)
            {
            default: MRRETURN(MATCH_NOMATCH);

            case 0x000d:
            if (eptr < md->end_subject && *eptr == 0x0a) eptr++;
            break;

            case 0x000a:
            break;

            case 0x000b:
            case 0x000c:
            case 0x0085:
            if (md->bsr_anycrlf) MRRETURN(MATCH_NOMATCH);
            break;
            }
          }
        break;

        case OP_NOT_HSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          switch(*eptr++)
            {
            default: break;
            case 0x09:      /* HT */
            case 0x20:      /* SPACE */
            case 0xa0:      /* NBSP */
            MRRETURN(MATCH_NOMATCH);
            }
          }
        break;

        case OP_HSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          switch(*eptr++)
            {
            default: MRRETURN(MATCH_NOMATCH);
            case 0x09:      /* HT */
            case 0x20:      /* SPACE */
            case 0xa0:      /* NBSP */
            break;
            }
          }
        break;

        case OP_NOT_VSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          switch(*eptr++)
            {
            default: break;
            case 0x0a:      /* LF */
            case 0x0b:      /* VT */
            case 0x0c:      /* FF */
            case 0x0d:      /* CR */
            case 0x85:      /* NEL */
            MRRETURN(MATCH_NOMATCH);
            }
          }
        break;

        case OP_VSPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          switch(*eptr++)
            {
            default: MRRETURN(MATCH_NOMATCH);
            case 0x0a:      /* LF */
            case 0x0b:      /* VT */
            case 0x0c:      /* FF */
            case 0x0d:      /* CR */
            case 0x85:      /* NEL */
            break;
            }
          }
        break;

        case OP_NOT_DIGIT:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if ((md->ctypes[*eptr++] & ctype_digit) != 0) MRRETURN(MATCH_NOMATCH);
          }
        break;

        case OP_DIGIT:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if ((md->ctypes[*eptr++] & ctype_digit) == 0) MRRETURN(MATCH_NOMATCH);
          }
        break;

        case OP_NOT_WHITESPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if ((md->ctypes[*eptr++] & ctype_space) != 0) MRRETURN(MATCH_NOMATCH);
          }
        break;

        case OP_WHITESPACE:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if ((md->ctypes[*eptr++] & ctype_space) == 0) MRRETURN(MATCH_NOMATCH);
          }
        break;

        case OP_NOT_WORDCHAR:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if ((md->ctypes[*eptr++] & ctype_word) != 0)
            MRRETURN(MATCH_NOMATCH);
          }
        break;

        case OP_WORDCHAR:
        for (i = 1; i <= min; i++)
          {
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if ((md->ctypes[*eptr++] & ctype_word) == 0)
            MRRETURN(MATCH_NOMATCH);
          }
        break;

        default:
        RRETURN(PCRE_ERROR_INTERNAL);
        }
      }

    /* If min = max, continue at the same level without recursing */

    if (min == max) continue;

    /* If minimizing, we have to test the rest of the pattern before each
    subsequent match. Again, separate the UTF-8 case for speed, and also
    separate the UCP cases. */

    if (minimize)
      {
#ifdef SUPPORT_UCP
      if (prop_type >= 0)
        {
        switch(prop_type)
          {
          case PT_ANY:
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM36);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if (prop_fail_result) MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_LAMP:
          for (fi = min;; fi++)
            {
            int chartype;
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM37);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            chartype = UCD_CHARTYPE(c);
            if ((chartype == ucp_Lu ||
                 chartype == ucp_Ll ||
                 chartype == ucp_Lt) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_GC:
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM38);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CATEGORY(c) == prop_value) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_PC:
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM39);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CHARTYPE(c) == prop_value) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_SC:
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM40);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_SCRIPT(c) == prop_value) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_ALNUM:
          for (fi = min;; fi++)
            {
            int category;
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM59);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            category = UCD_CATEGORY(c);
            if ((category == ucp_L || category == ucp_N) == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_SPACE:    /* Perl space */
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM60);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CATEGORY(c) == ucp_Z || c == CHAR_HT || c == CHAR_NL ||
                 c == CHAR_FF || c == CHAR_CR)
                   == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_PXSPACE:  /* POSIX space */
          for (fi = min;; fi++)
            {
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM61);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            if ((UCD_CATEGORY(c) == ucp_Z || c == CHAR_HT || c == CHAR_NL ||
                 c == CHAR_VT || c == CHAR_FF || c == CHAR_CR)
                   == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          case PT_WORD:
          for (fi = min;; fi++)
            {
            int category;
            RMATCH(eptr, ecode, offset_top, md, eptrb, RM62);
            if (rrc != MATCH_NOMATCH) RRETURN(rrc);
            if (fi >= max) MRRETURN(MATCH_NOMATCH);
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              MRRETURN(MATCH_NOMATCH);
              }
            TGETCHARINCTEST(c, eptr);
            category = UCD_CATEGORY(c);
            if ((category == ucp_L ||
                 category == ucp_N ||
                 c == CHAR_UNDERSCORE)
                   == prop_fail_result)
              MRRETURN(MATCH_NOMATCH);
            }
          /* Control never gets here */

          /* This should never occur */

          default:
          RRETURN(PCRE_ERROR_INTERNAL);
          }
        }

      /* Match extended Unicode sequences. We will get here only if the
      support is in the binary; otherwise a compile-time error occurs. */

      else if (ctype == OP_EXTUNI)
        {
        for (fi = min;; fi++)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM41);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (fi >= max) MRRETURN(MATCH_NOMATCH);
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          TGETCHARINCTEST(c, eptr);
          if (UCD_CATEGORY(c) == ucp_M) MRRETURN(MATCH_NOMATCH);
          while (eptr < md->end_subject)
            {
            int len = 1;
            if (!utf8) c = *eptr; else { TGETCHARLEN(c, eptr, len); }
            if (UCD_CATEGORY(c) != ucp_M) break;
            eptr += len;
            }
          }
        }
      else
#endif     /* SUPPORT_UCP */

#ifdef SUPPORT_UTF8
      /* UTF-8 mode */
      if (utf8)
        {
        for (fi = min;; fi++)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM42);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (fi >= max) MRRETURN(MATCH_NOMATCH);
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (ctype == OP_ANY && IS_NEWLINE(eptr))
            MRRETURN(MATCH_NOMATCH);
          TGETCHARINC(c, eptr);
          switch(ctype)
            {
            case OP_ANY:        /* This is the non-NL case */
            case OP_ALLANY:
            case OP_ANYBYTE:
            break;

            case OP_ANYNL:
            switch(c)
              {
              default: MRRETURN(MATCH_NOMATCH);
              case 0x000d:
              if (eptr < md->end_subject && *eptr == 0x0a) eptr++;
              break;
              case 0x000a:
              break;

              case 0x000b:
              case 0x000c:
              case 0x0085:
              case 0x2028:
              case 0x2029:
              if (md->bsr_anycrlf) MRRETURN(MATCH_NOMATCH);
              break;
              }
            break;

            case OP_NOT_HSPACE:
            switch(c)
              {
              default: break;
              case 0x09:      /* HT */
              case 0x20:      /* SPACE */
              case 0xa0:      /* NBSP */
              case 0x1680:    /* OGHAM SPACE MARK */
              case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
              case 0x2000:    /* EN QUAD */
              case 0x2001:    /* EM QUAD */
              case 0x2002:    /* EN SPACE */
              case 0x2003:    /* EM SPACE */
              case 0x2004:    /* THREE-PER-EM SPACE */
              case 0x2005:    /* FOUR-PER-EM SPACE */
              case 0x2006:    /* SIX-PER-EM SPACE */
              case 0x2007:    /* FIGURE SPACE */
              case 0x2008:    /* PUNCTUATION SPACE */
              case 0x2009:    /* THIN SPACE */
              case 0x200A:    /* HAIR SPACE */
              case 0x202f:    /* NARROW NO-BREAK SPACE */
              case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
              case 0x3000:    /* IDEOGRAPHIC SPACE */
              MRRETURN(MATCH_NOMATCH);
              }
            break;

            case OP_HSPACE:
            switch(c)
              {
              default: MRRETURN(MATCH_NOMATCH);
              case 0x09:      /* HT */
              case 0x20:      /* SPACE */
              case 0xa0:      /* NBSP */
              case 0x1680:    /* OGHAM SPACE MARK */
              case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
              case 0x2000:    /* EN QUAD */
              case 0x2001:    /* EM QUAD */
              case 0x2002:    /* EN SPACE */
              case 0x2003:    /* EM SPACE */
              case 0x2004:    /* THREE-PER-EM SPACE */
              case 0x2005:    /* FOUR-PER-EM SPACE */
              case 0x2006:    /* SIX-PER-EM SPACE */
              case 0x2007:    /* FIGURE SPACE */
              case 0x2008:    /* PUNCTUATION SPACE */
              case 0x2009:    /* THIN SPACE */
              case 0x200A:    /* HAIR SPACE */
              case 0x202f:    /* NARROW NO-BREAK SPACE */
              case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
              case 0x3000:    /* IDEOGRAPHIC SPACE */
              break;
              }
            break;

            case OP_NOT_VSPACE:
            switch(c)
              {
              default: break;
              case 0x0a:      /* LF */
              case 0x0b:      /* VT */
              case 0x0c:      /* FF */
              case 0x0d:      /* CR */
              case 0x85:      /* NEL */
              case 0x2028:    /* LINE SEPARATOR */
              case 0x2029:    /* PARAGRAPH SEPARATOR */
              MRRETURN(MATCH_NOMATCH);
              }
            break;

            case OP_VSPACE:
            switch(c)
              {
              default: MRRETURN(MATCH_NOMATCH);
              case 0x0a:      /* LF */
              case 0x0b:      /* VT */
              case 0x0c:      /* FF */
              case 0x0d:      /* CR */
              case 0x85:      /* NEL */
              case 0x2028:    /* LINE SEPARATOR */
              case 0x2029:    /* PARAGRAPH SEPARATOR */
              break;
              }
            break;

            case OP_NOT_DIGIT:
            if (c < 256 && (md->ctypes[c] & ctype_digit) != 0)
              MRRETURN(MATCH_NOMATCH);
            break;

            case OP_DIGIT:
            if (c >= 256 || (md->ctypes[c] & ctype_digit) == 0)
              MRRETURN(MATCH_NOMATCH);
            break;

            case OP_NOT_WHITESPACE:
            if (c < 256 && (md->ctypes[c] & ctype_space) != 0)
              MRRETURN(MATCH_NOMATCH);
            break;

            case OP_WHITESPACE:
            if  (c >= 256 || (md->ctypes[c] & ctype_space) == 0)
              MRRETURN(MATCH_NOMATCH);
            break;

            case OP_NOT_WORDCHAR:
            if (c < 256 && (md->ctypes[c] & ctype_word) != 0)
              MRRETURN(MATCH_NOMATCH);
            break;

            case OP_WORDCHAR:
            if (c >= 256 || (md->ctypes[c] & ctype_word) == 0)
              MRRETURN(MATCH_NOMATCH);
            break;

            default:
            RRETURN(PCRE_ERROR_INTERNAL);
            }
          }
        }
      else
#endif
      /* Not UTF-8 mode */
        {
        for (fi = min;; fi++)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM43);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (fi >= max) MRRETURN(MATCH_NOMATCH);
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            MRRETURN(MATCH_NOMATCH);
            }
          if (ctype == OP_ANY && IS_NEWLINE(eptr))
            MRRETURN(MATCH_NOMATCH);
          c = *eptr++;
          switch(ctype)
            {
            case OP_ANY:     /* This is the non-NL case */
            case OP_ALLANY:
            case OP_ANYBYTE:
            break;

            case OP_ANYNL:
            switch(c)
              {
              default: MRRETURN(MATCH_NOMATCH);
              case 0x000d:
              if (eptr < md->end_subject && *eptr == 0x0a) eptr++;
              break;

              case 0x000a:
              break;

              case 0x000b:
              case 0x000c:
              case 0x0085:
              if (md->bsr_anycrlf) MRRETURN(MATCH_NOMATCH);
              break;
              }
            break;

            case OP_NOT_HSPACE:
            switch(c)
              {
              default: break;
              case 0x09:      /* HT */
              case 0x20:      /* SPACE */
              case 0xa0:      /* NBSP */
              MRRETURN(MATCH_NOMATCH);
              }
            break;

            case OP_HSPACE:
            switch(c)
              {
              default: MRRETURN(MATCH_NOMATCH);
              case 0x09:      /* HT */
              case 0x20:      /* SPACE */
              case 0xa0:      /* NBSP */
              break;
              }
            break;

            case OP_NOT_VSPACE:
            switch(c)
              {
              default: break;
              case 0x0a:      /* LF */
              case 0x0b:      /* VT */
              case 0x0c:      /* FF */
              case 0x0d:      /* CR */
              case 0x85:      /* NEL */
              MRRETURN(MATCH_NOMATCH);
              }
            break;

            case OP_VSPACE:
            switch(c)
              {
              default: MRRETURN(MATCH_NOMATCH);
              case 0x0a:      /* LF */
              case 0x0b:      /* VT */
              case 0x0c:      /* FF */
              case 0x0d:      /* CR */
              case 0x85:      /* NEL */
              break;
              }
            break;

            case OP_NOT_DIGIT:
            if ((md->ctypes[c] & ctype_digit) != 0) MRRETURN(MATCH_NOMATCH);
            break;

            case OP_DIGIT:
            if ((md->ctypes[c] & ctype_digit) == 0) MRRETURN(MATCH_NOMATCH);
            break;

            case OP_NOT_WHITESPACE:
            if ((md->ctypes[c] & ctype_space) != 0) MRRETURN(MATCH_NOMATCH);
            break;

            case OP_WHITESPACE:
            if  ((md->ctypes[c] & ctype_space) == 0) MRRETURN(MATCH_NOMATCH);
            break;

            case OP_NOT_WORDCHAR:
            if ((md->ctypes[c] & ctype_word) != 0) MRRETURN(MATCH_NOMATCH);
            break;

            case OP_WORDCHAR:
            if ((md->ctypes[c] & ctype_word) == 0) MRRETURN(MATCH_NOMATCH);
            break;

            default:
            RRETURN(PCRE_ERROR_INTERNAL);
            }
          }
        }
      /* Control never gets here */
      }

    /* If maximizing, it is worth using inline code for speed, doing the type
    test once at the start (i.e. keep it out of the loop). Again, keep the
    UTF-8 and UCP stuff separate. */

    else
      {
      pp = eptr;  /* Remember where we started */

#ifdef SUPPORT_UCP
      if (prop_type >= 0)
        {
        switch(prop_type)
          {
          case PT_ANY:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            if (prop_fail_result) break;
            eptr+= len;
            }
          break;

          case PT_LAMP:
          for (i = min; i < max; i++)
            {
            int chartype;
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            chartype = UCD_CHARTYPE(c);
            if ((chartype == ucp_Lu ||
                 chartype == ucp_Ll ||
                 chartype == ucp_Lt) == prop_fail_result)
              break;
            eptr+= len;
            }
          break;

          case PT_GC:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            if ((UCD_CATEGORY(c) == prop_value) == prop_fail_result) break;
            eptr+= len;
            }
          break;

          case PT_PC:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            if ((UCD_CHARTYPE(c) == prop_value) == prop_fail_result) break;
            eptr+= len;
            }
          break;

          case PT_SC:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            if ((UCD_SCRIPT(c) == prop_value) == prop_fail_result) break;
            eptr+= len;
            }
          break;

          case PT_ALNUM:
          for (i = min; i < max; i++)
            {
            int category;
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            category = UCD_CATEGORY(c);
            if ((category == ucp_L || category == ucp_N) == prop_fail_result)
              break;
            eptr+= len;
            }
          break;

          case PT_SPACE:    /* Perl space */
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            if ((UCD_CATEGORY(c) == ucp_Z || c == CHAR_HT || c == CHAR_NL ||
                 c == CHAR_FF || c == CHAR_CR)
                 == prop_fail_result)
              break;
            eptr+= len;
            }
          break;

          case PT_PXSPACE:  /* POSIX space */
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            if ((UCD_CATEGORY(c) == ucp_Z || c == CHAR_HT || c == CHAR_NL ||
                 c == CHAR_VT || c == CHAR_FF || c == CHAR_CR)
                 == prop_fail_result)
              break;
            eptr+= len;
            }
          break;

          case PT_WORD:
          for (i = min; i < max; i++)
            {
            int category;
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLENTEST(c, eptr, len);
            category = UCD_CATEGORY(c);
            if ((category == ucp_L || category == ucp_N ||
                 c == CHAR_UNDERSCORE) == prop_fail_result)
              break;
            eptr+= len;
            }
          break;

          default:
          RRETURN(PCRE_ERROR_INTERNAL);
          }

        /* eptr is now past the end of the maximum run */

        if (possessive) continue;
        for(;;)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM44);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (eptr-- == pp) break;        /* Stop if tried at original pos */
          if (utf8) TBACKCHAR(eptr);
          }
        }

      /* Match extended Unicode sequences. We will get here only if the
      support is in the binary; otherwise a compile-time error occurs. */

      else if (ctype == OP_EXTUNI)
        {
        for (i = min; i < max; i++)
          {
          int len = 1;
          if (eptr >= md->end_subject)
            {
            SCHECK_PARTIAL();
            break;
            }
          if (!utf8) c = *eptr; else { TGETCHARLEN(c, eptr, len); }
          if (UCD_CATEGORY(c) == ucp_M) break;
          eptr += len;
          while (eptr < md->end_subject)
            {
            len = 1;
            if (!utf8) c = *eptr; else { TGETCHARLEN(c, eptr, len); }
            if (UCD_CATEGORY(c) != ucp_M) break;
            eptr += len;
            }
          }

        /* eptr is now past the end of the maximum run */

        if (possessive) continue;

        for(;;)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM45);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (eptr-- == pp) break;        /* Stop if tried at original pos */
          for (;;)                        /* Move back over one extended */
            {
            if (!utf8) c = *eptr; else
              {
              TBACKCHAR(eptr);
              TGETCHAR(c, eptr);
              }
            if (UCD_CATEGORY(c) != ucp_M) break;
            eptr--;
            }
          }
        }

      else
#endif   /* SUPPORT_UCP */

#ifdef SUPPORT_UTF8
      /* UTF-8 mode */

      if (utf8)
        {
        switch(ctype)
          {
          case OP_ANY:
          if (max < INT_MAX)
            {
            for (i = min; i < max; i++)
              {
              if (eptr >= md->end_subject)
                {
                SCHECK_PARTIAL();
                break;
                }
              if (IS_NEWLINE(eptr)) break;
              eptr++;
              TSKIPCHAR(eptr);
              }
            }

          /* Handle unlimited UTF-8 repeat */

          else
            {
            for (i = min; i < max; i++)
              {
              if (eptr >= md->end_subject)
                {
                SCHECK_PARTIAL();
                break;
                }
              if (IS_NEWLINE(eptr)) break;
              eptr++;
              TSKIPCHAR(eptr);
              }
            }
          break;

          case OP_ALLANY:
          if (max < INT_MAX)
            {
            for (i = min; i < max; i++)
              {
              if (eptr >= md->end_subject)
                {
                SCHECK_PARTIAL();
                break;
                }
              eptr++;
              TSKIPCHAR(eptr);
              }
            }
          else
            {
            eptr = md->end_subject;   /* Unlimited UTF-8 repeat */
            SCHECK_PARTIAL();
            }
          break;

          /* The byte case is the same as non-UTF8 */

          case OP_ANYBYTE:
          c = max - min;
          if (c > (unsigned int)(md->end_subject - eptr))
            {
            eptr = md->end_subject;
            SCHECK_PARTIAL();
            }
          else eptr += c;
          break;

          case OP_ANYNL:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c == 0x000d)
              {
              if (++eptr >= md->end_subject) break;
              if (*eptr == 0x000a) eptr++;
              }
            else
              {
              if (c != 0x000a &&
                  (md->bsr_anycrlf ||
                   (c != 0x000b && c != 0x000c &&
                    c != 0x0085 && c != 0x2028 && c != 0x2029)))
                break;
              eptr += len;
              }
            }
          break;

          case OP_NOT_HSPACE:
          case OP_HSPACE:
          for (i = min; i < max; i++)
            {
            BOOL gotspace;
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            switch(c)
              {
              default: gotspace = FALSE; break;
              case 0x09:      /* HT */
              case 0x20:      /* SPACE */
              case 0xa0:      /* NBSP */
              case 0x1680:    /* OGHAM SPACE MARK */
              case 0x180e:    /* MONGOLIAN VOWEL SEPARATOR */
              case 0x2000:    /* EN QUAD */
              case 0x2001:    /* EM QUAD */
              case 0x2002:    /* EN SPACE */
              case 0x2003:    /* EM SPACE */
              case 0x2004:    /* THREE-PER-EM SPACE */
              case 0x2005:    /* FOUR-PER-EM SPACE */
              case 0x2006:    /* SIX-PER-EM SPACE */
              case 0x2007:    /* FIGURE SPACE */
              case 0x2008:    /* PUNCTUATION SPACE */
              case 0x2009:    /* THIN SPACE */
              case 0x200A:    /* HAIR SPACE */
              case 0x202f:    /* NARROW NO-BREAK SPACE */
              case 0x205f:    /* MEDIUM MATHEMATICAL SPACE */
              case 0x3000:    /* IDEOGRAPHIC SPACE */
              gotspace = TRUE;
              break;
              }
            if (gotspace == (ctype == OP_NOT_HSPACE)) break;
            eptr += len;
            }
          break;

          case OP_NOT_VSPACE:
          case OP_VSPACE:
          for (i = min; i < max; i++)
            {
            BOOL gotspace;
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            switch(c)
              {
              default: gotspace = FALSE; break;
              case 0x0a:      /* LF */
              case 0x0b:      /* VT */
              case 0x0c:      /* FF */
              case 0x0d:      /* CR */
              case 0x85:      /* NEL */
              case 0x2028:    /* LINE SEPARATOR */
              case 0x2029:    /* PARAGRAPH SEPARATOR */
              gotspace = TRUE;
              break;
              }
            if (gotspace == (ctype == OP_NOT_VSPACE)) break;
            eptr += len;
            }
          break;

          case OP_NOT_DIGIT:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c < 256 && (md->ctypes[c] & ctype_digit) != 0) break;
            eptr+= len;
            }
          break;

          case OP_DIGIT:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c >= 256 ||(md->ctypes[c] & ctype_digit) == 0) break;
            eptr+= len;
            }
          break;

          case OP_NOT_WHITESPACE:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c < 256 && (md->ctypes[c] & ctype_space) != 0) break;
            eptr+= len;
            }
          break;

          case OP_WHITESPACE:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c >= 256 ||(md->ctypes[c] & ctype_space) == 0) break;
            eptr+= len;
            }
          break;

          case OP_NOT_WORDCHAR:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c < 256 && (md->ctypes[c] & ctype_word) != 0) break;
            eptr+= len;
            }
          break;

          case OP_WORDCHAR:
          for (i = min; i < max; i++)
            {
            int len = 1;
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            TGETCHARLEN(c, eptr, len);
            if (c >= 256 || (md->ctypes[c] & ctype_word) == 0) break;
            eptr+= len;
            }
          break;

          default:
          RRETURN(PCRE_ERROR_INTERNAL);
          }

        /* eptr is now past the end of the maximum run. If possessive, we are
        done (no backing up). Otherwise, match at this position; anything other
        than no match is immediately returned. For nomatch, back up one
        character, unless we are matching \R and the last thing matched was
        \r\n, in which case, back up two bytes. */

        if (possessive) continue;
        for(;;)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM46);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          if (eptr-- == pp) break;        /* Stop if tried at original pos */
          TBACKCHAR(eptr);
          if (ctype == OP_ANYNL && eptr > pp  && *eptr == '\n' &&
              eptr[-1] == '\r') eptr--;
          }
        }
      else
#endif  /* SUPPORT_UTF8 */

      /* Not UTF-8 mode */
        {
        switch(ctype)
          {
          case OP_ANY:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if (IS_NEWLINE(eptr)) break;
            eptr++;
            }
          break;

          case OP_ALLANY:
          case OP_ANYBYTE:
          c = max - min;
          if (c > (unsigned int)(md->end_subject - eptr))
            {
            eptr = md->end_subject;
            SCHECK_PARTIAL();
            }
          else eptr += c;
          break;

          case OP_ANYNL:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            c = *eptr;
            if (c == 0x000d)
              {
              if (++eptr >= md->end_subject) break;
              if (*eptr == 0x000a) eptr++;
              }
            else
              {
              if (c != 0x000a &&
                  (md->bsr_anycrlf ||
                    (c != 0x000b && c != 0x000c && c != 0x0085)))
                break;
              eptr++;
              }
            }
          break;

          case OP_NOT_HSPACE:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            c = *eptr;
            if (c == 0x09 || c == 0x20 || c == 0xa0) break;
            eptr++;
            }
          break;

          case OP_HSPACE:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            c = *eptr;
            if (c != 0x09 && c != 0x20 && c != 0xa0) break;
            eptr++;
            }
          break;

          case OP_NOT_VSPACE:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            c = *eptr;
            if (c == 0x0a || c == 0x0b || c == 0x0c || c == 0x0d || c == 0x85)
              break;
            eptr++;
            }
          break;

          case OP_VSPACE:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            c = *eptr;
            if (c != 0x0a && c != 0x0b && c != 0x0c && c != 0x0d && c != 0x85)
              break;
            eptr++;
            }
          break;

          case OP_NOT_DIGIT:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if ((md->ctypes[*eptr] & ctype_digit) != 0) break;
            eptr++;
            }
          break;

          case OP_DIGIT:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if ((md->ctypes[*eptr] & ctype_digit) == 0) break;
            eptr++;
            }
          break;

          case OP_NOT_WHITESPACE:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if ((md->ctypes[*eptr] & ctype_space) != 0) break;
            eptr++;
            }
          break;

          case OP_WHITESPACE:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if ((md->ctypes[*eptr] & ctype_space) == 0) break;
            eptr++;
            }
          break;

          case OP_NOT_WORDCHAR:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if ((md->ctypes[*eptr] & ctype_word) != 0) break;
            eptr++;
            }
          break;

          case OP_WORDCHAR:
          for (i = min; i < max; i++)
            {
            if (eptr >= md->end_subject)
              {
              SCHECK_PARTIAL();
              break;
              }
            if ((md->ctypes[*eptr] & ctype_word) == 0) break;
            eptr++;
            }
          break;

          default:
          RRETURN(PCRE_ERROR_INTERNAL);
          }

        /* eptr is now past the end of the maximum run. If possessive, we are
        done (no backing up). Otherwise, match at this position; anything other
        than no match is immediately returned. For nomatch, back up one
        character (byte), unless we are matching \R and the last thing matched
        was \r\n, in which case, back up two bytes. */

        if (possessive) continue;
        while (eptr >= pp)
          {
          RMATCH(eptr, ecode, offset_top, md, eptrb, RM47);
          if (rrc != MATCH_NOMATCH) RRETURN(rrc);
          eptr--;
          if (ctype == OP_ANYNL && eptr > pp  && *eptr == '\n' &&
              eptr[-1] == '\r') eptr--;
          }
        }

      /* Get here if we can't make it match with any permitted repetitions */

      MRRETURN(MATCH_NOMATCH);
      }
    /* Control never gets here */

    /* There's been some horrible disaster. Arrival here can only mean there is
    something seriously wrong in the code above or the OP_xxx definitions. */

    default:
    DPRINTF(("Unknown opcode %d\n", *ecode));
    RRETURN(PCRE_ERROR_UNKNOWN_OPCODE);
    }

  /* Do not stick any code in here without much thought; it is assumed
  that "continue" in the code above comes out to here to repeat the main
  loop. */

  }             /* End of main loop */
/* Control never reaches here */


/* When compiling to use the heap rather than the stack for recursive calls to
match(), the RRETURN() macro jumps here. The number that is saved in
frame->Xwhere indicates which label we actually want to return to. */

#ifdef NO_RECURSE
#define LBL(val) case val: goto L_RM##val;
HEAP_RETURN:
switch (frame->Xwhere)
  {
  LBL( 1) LBL( 2) LBL( 3) LBL( 4) LBL( 5) LBL( 6) LBL( 7) LBL( 8)
  LBL( 9) LBL(10) LBL(11) LBL(12) LBL(13) LBL(14) LBL(15) LBL(17)
  LBL(19) LBL(24) LBL(25) LBL(26) LBL(27) LBL(29) LBL(31) LBL(33)
  LBL(35) LBL(43) LBL(47) LBL(48) LBL(49) LBL(50) LBL(51) LBL(52)
  LBL(53) LBL(54) LBL(55) LBL(56) LBL(57) LBL(58) LBL(63)
#ifdef SUPPORT_UTF8
  LBL(16) LBL(18) LBL(20) LBL(21) LBL(22) LBL(23) LBL(28) LBL(30)
  LBL(32) LBL(34) LBL(42) LBL(46)
#ifdef SUPPORT_UCP
  LBL(36) LBL(37) LBL(38) LBL(39) LBL(40) LBL(41) LBL(44) LBL(45)
  LBL(59) LBL(60) LBL(61) LBL(62)
#endif  /* SUPPORT_UCP */
#endif  /* SUPPORT_UTF8 */
  default:
  DPRINTF(("jump error in pcre match: label %d non-existent\n", frame->Xwhere));
  return PCRE_ERROR_INTERNAL;
  }
#undef LBL
#endif  /* NO_RECURSE */
}


/***************************************************************************
****************************************************************************
                   RECURSION IN THE match() FUNCTION

Undefine all the macros that were defined above to handle this. */

#ifdef NO_RECURSE
#undef eptr
#undef ecode
#undef mstart
#undef offset_top
#undef eptrb
#undef flags

#undef callpat
#undef charptr
#undef data
#undef next
#undef pp
#undef prev
#undef saved_eptr

#undef new_recursive

#undef cur_is_word
#undef condition
#undef prev_is_word

#undef ctype
#undef length
#undef max
#undef min
#undef number
#undef offset
#undef op
#undef save_capture_last
#undef save_offset1
#undef save_offset2
#undef save_offset3
#undef stacksave

#undef newptrb

#endif

/* These two are defined as macros in both cases */

#undef fc
#undef fi

/***************************************************************************
***************************************************************************/



/*************************************************
*         Execute a Regular Expression           *
*************************************************/

/* This function applies a compiled re to a subject string and picks out
portions of the string if it matches. Two elements in the vector are set for
each substring: the offsets to the start and end of the substring.

Arguments:
  argument_re     points to the compiled expression
  extra_data      points to extra data or is NULL
  subject         points to the subject string
  length          length of subject string (may contain binary zeros)
  start_offset    where to start in the subject string
  options         option bits
  offsets         points to a vector of ints to be filled in with offsets
  offsetcount     the number of elements in the vector

Returns:          > 0 => success; value is the number of elements filled in
                  = 0 => success, but offsets is not big enough
                   -1 => failed to match
                 < -1 => some kind of unexpected problem
*/

PCRE_EXP_DEFN int PCRE_CALL_CONVENTION
pcre_exec(const pcre *argument_re, const pcre_extra *extra_data,
  PCRE_SPTR subject, int length, int start_offset, int options, int *offsets,
  int offsetcount)
{
int rc, ocount;
int first_byte = -1;
int req_byte = -1;
int req_byte2 = -1;
int newline;
BOOL using_temporary_offsets = FALSE;
BOOL anchored;
BOOL startline;
BOOL firstline;
BOOL first_byte_caseless = FALSE;
BOOL req_byte_caseless = FALSE;
#ifdef SUPPORT_UTF8 /* AutoHotkey: This helps detected unintended usages of utf8. */
#ifdef SUPPORT_UTF8_ONLY
    const BOOL utf8 = TRUE;
#else
    BOOL utf8;
#endif
#endif /* AutoHotkey. */
match_data match_block;
match_data *md = &match_block;
const uschar *tables;
const uschar *start_bits = NULL;
USPTR start_match = (USPTR)subject + start_offset;
USPTR end_subject;
USPTR start_partial = NULL;
USPTR req_byte_ptr = start_match - 1;

pcre_study_data internal_study;
const pcre_study_data *study;

real_pcre internal_re;
const real_pcre *external_re = (const real_pcre *)argument_re;
const real_pcre *re = external_re;

/* Plausibility checks */

if ((options & ~PUBLIC_EXEC_OPTIONS) != 0) return PCRE_ERROR_BADOPTION;
if (re == NULL || subject == NULL ||
   (offsets == NULL && offsetcount > 0)) return PCRE_ERROR_NULL;
if (offsetcount < 0) return PCRE_ERROR_BADCOUNT;
if (start_offset < 0 || start_offset > length) return PCRE_ERROR_BADOFFSET;

/* This information is for finding all the numbers associated with a given
name, for condition testing. */

md->name_table = (utchar *)((uschar *)re + re->name_table_offset);
md->name_count = re->name_count;
md->name_entry_size = re->name_entry_size;

/* Fish out the optional data from the extra_data structure, first setting
the default values. */

study = NULL;
md->match_limit = MATCH_LIMIT;
md->match_limit_recursion = MATCH_LIMIT_RECURSION;
md->callout_data = NULL;

/* The table pointer is always in native byte order. */

tables = external_re->tables;

if (extra_data != NULL)
  {
  register unsigned int flags = extra_data->flags;
  if ((flags & PCRE_EXTRA_STUDY_DATA) != 0)
    study = (const pcre_study_data *)extra_data->study_data;
  if ((flags & PCRE_EXTRA_MATCH_LIMIT) != 0)
    md->match_limit = extra_data->match_limit;
  if ((flags & PCRE_EXTRA_MATCH_LIMIT_RECURSION) != 0)
    md->match_limit_recursion = extra_data->match_limit_recursion;
  if ((flags & PCRE_EXTRA_CALLOUT_DATA) != 0)
    md->callout_data = extra_data->callout_data;
  if ((flags & PCRE_EXTRA_TABLES) != 0) tables = extra_data->tables;
  }

/* If the exec call supplied NULL for tables, use the inbuilt ones. This
is a feature that makes it possible to save compiled regex and re-use them
in other programs later. */

if (tables == NULL) tables = _pcre_default_tables;

/* Check that the first field in the block is the magic number. If it is not,
test for a regex that was compiled on a host of opposite endianness. If this is
the case, flipped values are put in internal_re and internal_study if there was
study data too. */

if (re->magic_number != MAGIC_NUMBER)
  {
  re = _pcre_try_flipped(re, &internal_re, study, &internal_study);
  if (re == NULL) return PCRE_ERROR_BADMAGIC;
  if (study != NULL) study = &internal_study;
  }

/* Set up other data */

anchored = ((re->options | options) & PCRE_ANCHORED) != 0;
startline = (re->flags & PCRE_STARTLINE) != 0;
firstline = (re->options & PCRE_FIRSTLINE) != 0;

/* The code starts after the real_pcre block and the capture name table. */

md->start_code = (const uschar *)external_re + re->name_table_offset +
  (re->name_count * re->name_entry_size * sizeof(utchar));

md->start_subject = (USPTR)subject;
md->start_offset = start_offset;
md->end_subject = md->start_subject + length;
end_subject = md->end_subject;

md->endonly = (re->options & PCRE_DOLLAR_ENDONLY) != 0;
#ifdef SUPPORT_UTF8 /* AutoHotkey. */
#ifndef SUPPORT_UTF8_ONLY
    utf8 = md->utf8 = (re->options & PCRE_UTF8) != 0;
#else
    md->utf8 = TRUE;
#endif
    md->use_ucp = (re->options & PCRE_UCP) != 0;
#endif /* AutoHotkey. */
md->jscript_compat = (re->options & PCRE_JAVASCRIPT_COMPAT) != 0;

/* Some options are unpacked into BOOL variables in the hope that testing
them will be faster than individual option bits. */

md->notbol = (options & PCRE_NOTBOL) != 0;
md->noteol = (options & PCRE_NOTEOL) != 0;
md->notempty = (options & PCRE_NOTEMPTY) != 0;
md->notempty_atstart = (options & PCRE_NOTEMPTY_ATSTART) != 0;
md->partial = ((options & PCRE_PARTIAL_HARD) != 0)? 2 :
              ((options & PCRE_PARTIAL_SOFT) != 0)? 1 : 0;


md->hitend = FALSE;
md->mark = NULL;                        /* In case never set */

md->recursive = NULL;                   /* No recursion at top level */

md->lcc = tables + lcc_offset;
md->ctypes = tables + ctypes_offset;

/* Handle different \R options. */

switch (options & (PCRE_BSR_ANYCRLF|PCRE_BSR_UNICODE))
  {
  case 0:
  if ((re->options & (PCRE_BSR_ANYCRLF|PCRE_BSR_UNICODE)) != 0)
    md->bsr_anycrlf = (re->options & PCRE_BSR_ANYCRLF) != 0;
  else
#ifdef BSR_ANYCRLF
  md->bsr_anycrlf = TRUE;
#else
  md->bsr_anycrlf = FALSE;
#endif
  break;

  case PCRE_BSR_ANYCRLF:
  md->bsr_anycrlf = TRUE;
  break;

  case PCRE_BSR_UNICODE:
  md->bsr_anycrlf = FALSE;
  break;

  default: return PCRE_ERROR_BADNEWLINE;
  }

/* Handle different types of newline. The three bits give eight cases. If
nothing is set at run time, whatever was used at compile time applies. */

switch ((((options & PCRE_NEWLINE_BITS) == 0)? re->options :
        (pcre_uint32)options) & PCRE_NEWLINE_BITS)
  {
  case 0: newline = NEWLINE; break;   /* Compile-time default */
  case PCRE_NEWLINE_CR: newline = CHAR_CR; break;
  case PCRE_NEWLINE_LF: newline = CHAR_NL; break;
  case PCRE_NEWLINE_CR+
       PCRE_NEWLINE_LF: newline = (CHAR_CR << 8) | CHAR_NL; break;
  case PCRE_NEWLINE_ANY: newline = -1; break;
  case PCRE_NEWLINE_ANYCRLF: newline = -2; break;
  default: return PCRE_ERROR_BADNEWLINE;
  }

if (newline == -2)
  {
  md->nltype = NLTYPE_ANYCRLF;
  }
else if (newline < 0)
  {
  md->nltype = NLTYPE_ANY;
  }
else
  {
  md->nltype = NLTYPE_FIXED;
  if (newline > 255)
    {
    md->nllen = 2;
    md->nl[0] = (newline >> 8) & 255;
    md->nl[1] = newline & 255;
    }
  else
    {
    md->nllen = 1;
    md->nl[0] = newline;
    }
  }

/* Partial matching was originally supported only for a restricted set of
regexes; from release 8.00 there are no restrictions, but the bits are still
defined (though never set). So there's no harm in leaving this code. */

if (md->partial && (re->flags & PCRE_NOPARTIAL) != 0)
  return PCRE_ERROR_BADPARTIAL;

/* Check a UTF-8 or UTF-16 string if required. Pass back the character offset
and error code for an invalid string if a results vector is available. */

#ifdef SUPPORT_UTF8
if (utf8 && (options & PCRE_NO_UTF8_CHECK) == 0)
  {
  int erroroffset;
  int errorcode = _pcre_valid_utf((USPTR)subject, length, &erroroffset);
  if (errorcode != 0)
    {
    if (offsetcount >= 2)
      {
      offsets[0] = erroroffset;
      offsets[1] = errorcode;
      }
    return (errorcode <= PCRE_UTF8_ERR5 && md->partial > 1)?
      PCRE_ERROR_SHORTUTF8 : PCRE_ERROR_BADUTF8;
    }

  /* Check that a start_offset points to the start of a UTF-8 character. */

  if (start_offset > 0 && start_offset < length &&
#ifdef PCRE_USE_UTF16
      IS_SURROGATE_PAIR(subject[start_offset-1], subject[start_offset]))
#else
      (((USPTR)subject)[start_offset] & 0xc0) == 0x80)
#endif
    return PCRE_ERROR_BADUTF8_OFFSET;
  }
#endif

/* If the expression has got more back references than the offsets supplied can
hold, we get a temporary chunk of working store to use during the matching.
Otherwise, we can use the vector supplied, rounding down its size to a multiple
of 3. */

ocount = offsetcount - (offsetcount % 3);

if (re->top_backref > 0 && re->top_backref >= ocount/3)
  {
  ocount = re->top_backref * 3 + 3;
  md->offset_vector = (int *)(pcre_malloc)(ocount * sizeof(int));
  if (md->offset_vector == NULL) return PCRE_ERROR_NOMEMORY;
  using_temporary_offsets = TRUE;
  DPRINTF(("Got memory to hold back references\n"));
  }
else md->offset_vector = offsets;

md->offset_end = ocount;
md->offset_max = (2*ocount)/3;
md->offset_overflow = FALSE;
md->capture_last = -1;

/* Reset the working variable associated with each extraction. These should
never be used unless previously set, but they get saved and restored, and so we
initialize them to avoid reading uninitialized locations. Also, unset the
offsets for the matched string. This is really just for tidiness with callouts,
in case they inspect these fields. */

if (md->offset_vector != NULL)
  {
  register int *iptr = md->offset_vector + ocount;
  register int *iend = iptr - re->top_bracket;
  if (iend < md->offset_vector + 2) iend = md->offset_vector + 2;
  while (--iptr >= iend) *iptr = -1;
  md->offset_vector[0] = md->offset_vector[1] = -1;
  }

/* Set up the first character to match, if available. The first_byte value is
never set for an anchored regular expression, but the anchoring may be forced
at run time, so we have to test for anchoring. The first char may be unset for
an unanchored pattern, of course. If there's no first char and the pattern was
studied, there may be a bitmap of possible first characters. */

if (!anchored)
  {
  if ((re->flags & PCRE_FIRSTSET) != 0)
    {
    first_byte = re->first_byte & 255;
    if ((first_byte_caseless = ((re->first_byte & REQ_CASELESS) != 0)) == TRUE)
      first_byte = md->lcc[first_byte];
    }
  else
    if (!startline && study != NULL &&
      (study->flags & PCRE_STUDY_MAPPED) != 0)
        start_bits = study->start_bits;
  }

/* For anchored or unanchored matches, there may be a "last known required
character" set. */

if ((re->flags & PCRE_REQCHSET) != 0)
  {
  req_byte = re->req_byte & 255;
  req_byte_caseless = (re->req_byte & REQ_CASELESS) != 0;
  req_byte2 = (tables + fcc_offset)[req_byte];  /* case flipped */
  }




/* ==========================================================================*/

/* Loop for handling unanchored repeated matching attempts; for anchored regexs
the loop runs just once. */

for(;;)
  {
  USPTR save_end_subject = end_subject;
  USPTR new_start_match;

  /* If firstline is TRUE, the start of the match is constrained to the first
  line of a multiline string. That is, the match must be before or at the first
  newline. Implement this by temporarily adjusting end_subject so that we stop
  scanning at a newline. If the match fails at the newline, later code breaks
  this loop. */

  if (firstline)
    {
    USPTR t = start_match;
#ifdef SUPPORT_UTF8_SUBJECT
    if (utf8)
      {
      while (t < md->end_subject && !IS_NEWLINE(t))
        {
        t++;
        while (t < end_subject && (*t & 0xc0) == 0x80) t++;
        }
      }
    else
#endif
    while (t < md->end_subject && !IS_NEWLINE(t)) t++;
    end_subject = t;
    }

  /* There are some optimizations that avoid running the match if a known
  starting point is not found, or if a known later character is not present.
  However, there is an option that disables these, for testing and for ensuring
  that all callouts do actually occur. The option can be set in the regex by
  (*NO_START_OPT) or passed in match-time options. */

  if (((options | re->options) & PCRE_NO_START_OPTIMIZE) == 0)
    {
    /* Advance to a unique first byte if there is one. */

    if (first_byte >= 0)
      {
      if (first_byte_caseless)
        while (start_match < end_subject && md->lcc[*start_match] != first_byte)
          start_match++;
      else
        while (start_match < end_subject && *start_match != first_byte)
          start_match++;
      }

    /* Or to just after a linebreak for a multiline match */

    else if (startline)
      {
      if (start_match > md->start_subject + start_offset)
        {
#ifdef SUPPORT_UTF8
        if (utf8)
          {
          while (start_match < end_subject && !WAS_NEWLINE(start_match))
            {
            start_match++;
            TSKIPCHAR(start_match);
            }
          }
        else
#endif
        while (start_match < end_subject && !WAS_NEWLINE(start_match))
          start_match++;

        /* If we have just passed a CR and the newline option is ANY or ANYCRLF,
        and we are now at a LF, advance the match position by one more character.
        */

        if (start_match[-1] == CHAR_CR &&
             (md->nltype == NLTYPE_ANY || md->nltype == NLTYPE_ANYCRLF) &&
             start_match < end_subject &&
             *start_match == CHAR_NL)
          start_match++;
        }
      }

    /* Or to a non-unique first byte after study */

    else if (start_bits != NULL)
      {
#ifdef PCRE_USE_UTF16
      /* AutoHotkey: See toward the end of pcre_study() for comments. */
      BOOL ascii_only = (start_bits[31] != 0);
#endif
      while (start_match < end_subject)
        {
        register unsigned int c = *start_match;
#ifdef PCRE_USE_UTF16
        if (c > 127 ? ascii_only : (start_bits[c/8] & (1 << (c&7))) == 0)
#else
        if ((start_bits[c/8] & (1 << (c&7))) == 0)
#endif
          {
          start_match++;
#ifdef SUPPORT_UTF8
          if (utf8) TSKIPCHAR(start_match);
#endif
          }
        else break;
        }
      }
    }   /* Starting optimizations */

  /* Restore fudged end_subject */

  end_subject = save_end_subject;

  /* The following two optimizations are disabled for partial matching or if
  disabling is explicitly requested. */

  if ((options & PCRE_NO_START_OPTIMIZE) == 0 && !md->partial)
    {
    /* If the pattern was studied, a minimum subject length may be set. This is
    a lower bound; no actual string of that length may actually match the
    pattern. Although the value is, strictly, in characters, we treat it as
    bytes to avoid spending too much time in this optimization. */

    if (study != NULL && (study->flags & PCRE_STUDY_MINLEN) != 0 &&
        (pcre_uint32)(end_subject - start_match) < study->minlength)
      {
      rc = MATCH_NOMATCH;
      break;
      }

    /* If req_byte is set, we know that that character must appear in the
    subject for the match to succeed. If the first character is set, req_byte
    must be later in the subject; otherwise the test starts at the match point.
    This optimization can save a huge amount of backtracking in patterns with
    nested unlimited repeats that aren't going to match. Writing separate code
    for cased/caseless versions makes it go faster, as does using an
    autoincrement and backing off on a match.

    HOWEVER: when the subject string is very, very long, searching to its end
    can take a long time, and give bad performance on quite ordinary patterns.
    This showed up when somebody was matching something like /^\d+C/ on a
    32-megabyte string... so we don't do this when the string is sufficiently
    long. */

    if (req_byte >= 0 && end_subject - start_match < REQ_BYTE_MAX)
      {
      register USPTR p = start_match + ((first_byte >= 0)? 1 : 0);

      /* We don't need to repeat the search if we haven't yet reached the
      place we found it at last time. */

      if (p > req_byte_ptr)
        {
        if (req_byte_caseless)
          {
          while (p < end_subject)
            {
            register int pp = *p++;
            if (pp == req_byte || pp == req_byte2) { p--; break; }
            }
          }
        else
          {
          while (p < end_subject)
            {
            if (*p++ == req_byte) { p--; break; }
            }
          }

        /* If we can't find the required character, break the matching loop,
        forcing a match failure. */

        if (p >= end_subject)
          {
          rc = MATCH_NOMATCH;
          break;
          }

        /* If we have found the required character, save the point where we
        found it, so that we don't search again next time round the loop if
        the start hasn't passed this character yet. */

        req_byte_ptr = p;
        }
      }
    }

#ifdef PCRE_DEBUG  /* Sigh. Some compilers never learn. */
  printf(">>>> Match against: ");
  pchars(start_match, end_subject - start_match, TRUE, md);
  printf("\n");
#endif

  /* OK, we can now run the match. If "hitend" is set afterwards, remember the
  first starting point for which a partial match was found. */

  md->start_match_ptr = start_match;
  md->start_used_ptr = start_match;
  md->match_call_count = 0;
  md->match_function_type = 0;
  md->end_offset_top = 0;
  rc = match(start_match, md->start_code, start_match, NULL, 2, md, NULL, 0);
  if (md->hitend && start_partial == NULL) start_partial = md->start_used_ptr;

  switch(rc)
    {
    /* SKIP passes back the next starting point explicitly, but if it is the
    same as the match we have just done, treat it as NOMATCH. */

    case MATCH_SKIP:
    if (md->start_match_ptr != start_match)
      {
      new_start_match = md->start_match_ptr;
      break;
      }
    /* Fall through */

    /* If MATCH_SKIP_ARG reaches this level it means that a MARK that matched
    the SKIP's arg was not found. We also treat this as NOMATCH. */

    case MATCH_SKIP_ARG:
    /* Fall through */

    /* NOMATCH and PRUNE advance by one character. THEN at this level acts
    exactly like PRUNE. */

    case MATCH_NOMATCH:
    case MATCH_PRUNE:
    case MATCH_THEN:
    new_start_match = start_match + 1;
#ifdef SUPPORT_UTF8
    if (utf8) TSKIPCHAR(new_start_match);
#endif
    break;

    /* COMMIT disables the bumpalong, but otherwise behaves as NOMATCH. */

    case MATCH_COMMIT:
    rc = MATCH_NOMATCH;
    goto ENDLOOP;

    /* Any other return is either a match, or some kind of error. */

    default:
    goto ENDLOOP;
    }

  /* Control reaches here for the various types of "no match at this point"
  result. Reset the code to MATCH_NOMATCH for subsequent checking. */

  rc = MATCH_NOMATCH;

  /* If PCRE_FIRSTLINE is set, the match must happen before or at the first
  newline in the subject (though it may continue over the newline). Therefore,
  if we have just failed to match, starting at a newline, do not continue. */

  if (firstline && IS_NEWLINE(start_match)) break;

  /* Advance to new matching position */

  start_match = new_start_match;

  /* Break the loop if the pattern is anchored or if we have passed the end of
  the subject. */

  if (anchored || start_match > end_subject) break;

  /* If we have just passed a CR and we are now at a LF, and the pattern does
  not contain any explicit matches for \r or \n, and the newline option is CRLF
  or ANY or ANYCRLF, advance the match position by one more character. */

  if (start_match[-1] == CHAR_CR &&
      start_match < end_subject &&
      *start_match == CHAR_NL &&
      (re->flags & PCRE_HASCRORLF) == 0 &&
        (md->nltype == NLTYPE_ANY ||
         md->nltype == NLTYPE_ANYCRLF ||
         md->nllen == 2))
    start_match++;

  md->mark = NULL;   /* Reset for start of next match attempt */
  }                  /* End of for(;;) "bumpalong" loop */

/* ==========================================================================*/

/* We reach here when rc is not MATCH_NOMATCH, or if one of the stopping
conditions is true:

(1) The pattern is anchored or the match was failed by (*COMMIT);

(2) We are past the end of the subject;

(3) PCRE_FIRSTLINE is set and we have failed to match at a newline, because
    this option requests that a match occur at or before the first newline in
    the subject.

When we have a match and the offset vector is big enough to deal with any
backreferences, captured substring offsets will already be set up. In the case
where we had to get some local store to hold offsets for backreference
processing, copy those that we can. In this case there need not be overflow if
certain parts of the pattern were not used, even though there are more
capturing parentheses than vector slots. */

ENDLOOP:

if (rc == MATCH_MATCH || rc == MATCH_ACCEPT)
  {
  if (using_temporary_offsets)
    {
    if (offsetcount >= 4)
      {
      memcpy(offsets + 2, md->offset_vector + 2,
        (offsetcount - 2) * sizeof(int));
      DPRINTF(("Copied offsets from temporary memory\n"));
      }
    if (md->end_offset_top > offsetcount) md->offset_overflow = TRUE;
    DPRINTF(("Freeing temporary memory\n"));
    (pcre_free)(md->offset_vector);
    }

  /* Set the return code to the number of captured strings, or 0 if there are
  too many to fit into the vector. */

  rc = md->offset_overflow? 0 : md->end_offset_top/2;

  /* If there is space in the offset vector, set any unused pairs at the end of
  the pattern to -1 for backwards compatibility. It is documented that this
  happens. In earlier versions, the whole set of potential capturing offsets
  was set to -1 each time round the loop, but this is handled differently now.
  "Gaps" are set to -1 dynamically instead (this fixes a bug). Thus, it is only
  those at the end that need unsetting here. We can't just unset them all at
  the start of the whole thing because they may get set in one branch that is
  not the final matching branch. */

  if (md->end_offset_top/2 <= re->top_bracket && offsets != NULL)
    {
    register int *iptr, *iend;
    int resetcount = 2 + re->top_bracket * 2;
    if (resetcount > offsetcount) resetcount = ocount;
    iptr = offsets + md->end_offset_top;
    iend = offsets + resetcount;
    while (iptr < iend) *iptr++ = -1;
    }

  /* If there is space, set up the whole thing as substring 0. The value of
  md->start_match_ptr might be modified if \K was encountered on the success
  matching path. */

  if (offsetcount < 2) rc = 0; else
    {
    offsets[0] = (int)(md->start_match_ptr - md->start_subject);
    offsets[1] = (int)(md->end_match_ptr - md->start_subject);
    }

  DPRINTF((">>>> returning %d\n", rc));
  goto RETURN_MARK;
  }

/* Control gets here if there has been an error, or if the overall match
attempt has failed at all permitted starting positions. */

if (using_temporary_offsets)
  {
  DPRINTF(("Freeing temporary memory\n"));
  (pcre_free)(md->offset_vector);
  }

/* For anything other than nomatch or partial match, just return the code. */

if (rc != MATCH_NOMATCH && rc != PCRE_ERROR_PARTIAL)
  {
  DPRINTF((">>>> error: returning %d\n", rc));
  return rc;
  }

/* Handle partial matches - disable any mark data */

if (start_partial != NULL)
  {
  DPRINTF((">>>> returning PCRE_ERROR_PARTIAL\n"));
  md->mark = NULL;
  if (offsetcount > 1)
    {
    offsets[0] = (int)(start_partial - (USPTR)subject);
    offsets[1] = (int)(end_subject - (USPTR)subject);
    }
  rc = PCRE_ERROR_PARTIAL;
  }

/* This is the classic nomatch case */

else
  {
  DPRINTF((">>>> returning PCRE_ERROR_NOMATCH\n"));
  rc = PCRE_ERROR_NOMATCH;
  }

/* Return the MARK data if it has been requested. */

RETURN_MARK:

if (extra_data != NULL && (extra_data->flags & PCRE_EXTRA_MARK) != 0)
  *(extra_data->mark) = (unsigned char *)(md->mark);
return rc;
}

/* End of pcre_exec.c */
