
#ifdef UNICODE

#define pcret								pcre16
#define pcret_extra							pcre16_extra
#define pcret_callout_block					pcre16_callout_block
#define pcret_malloc						pcre16_malloc
#define pcret_free							pcre16_free
#define pcret_stack_malloc					pcre16_stack_malloc
#define pcret_stack_free					pcre16_stack_free
#define pcret_callout						pcre16_callout

#define pcret_resolve_user_callout			pcre16_resolve_user_callout

#define pcret_compile						pcre16_compile
#define pcret_compile2						pcre16_compile2
#define pcret_config						pcre16_config
#define pcret_copy_named_substring			pcre16_copy_named_substring
#define pcret_copy_substring				pcre16_copy_substring
#define pcret_dfa_exec						pcre16_dfa_exec
#define pcret_exec							pcre16_exec
#define pcret_free_substring				pcre16_free_substring
#define pcret_free_substring_list			pcre16_free_substring_list
#define pcret_fullinfo						pcre16_fullinfo
#define pcret_get_named_substring			pcre16_get_named_substring
#define pcret_get_stringnumber				pcre16_get_stringnumber
#define pcret_get_stringtable_entries		pcre16_get_stringtable_entries
#define pcret_get_substring					pcre16_get_substring
#define pcret_get_substring_list			pcre16_get_substring_list
#define pcret_maketables					pcre16_maketables
#define pcret_refcount						pcre16_refcount
#define pcret_study							pcre16_study
#define pcret_free_study					pcre16_free_study
#define pcret_version						pcre16_version
#define pcret_pattern_to_host_byte_order	pcre16_pattern_to_host_byte_order
#define pcret_jit_stack_alloc				pcre16_jit_stack_alloc
#define pcret_jit_stack_free				pcre16_jit_stack_free
#define pcret_assign_jit_stack				pcre16_assign_jit_stack

#else

#define pcret								pcre
#define pcret_extra							pcre_extra
#define pcret_callout_block					pcre_callout_block
#define pcret_malloc						pcre_malloc
#define pcret_free							pcre_free
#define pcret_stack_malloc					pcre_stack_malloc
#define pcret_stack_free					pcre_stack_free
#define pcret_callout						pcre_callout

#define pcret_resolve_user_callout			pcre_resolve_user_callout

#define pcret_compile						pcre_compile
#define pcret_compile2						pcre_compile2
#define pcret_config						pcre_config
#define pcret_copy_named_substring			pcre_copy_named_substring
#define pcret_copy_substring				pcre_copy_substring
#define pcret_dfa_exec						pcre_dfa_exec
#define pcret_exec							pcre_exec
#define pcret_free_substring				pcre_free_substring
#define pcret_free_substring_list			pcre_free_substring_list
#define pcret_fullinfo						pcre_fullinfo
#define pcret_get_named_substring			pcre_get_named_substring
#define pcret_get_stringnumber				pcre_get_stringnumber
#define pcret_get_stringtable_entries		pcre_get_stringtable_entries
#define pcret_get_substring					pcre_get_substring
#define pcret_get_substring_list			pcre_get_substring_list
#define pcret_maketables					pcre_maketables
#define pcret_refcount						pcre_refcount
#define pcret_study							pcre_study
#define pcret_free_study					pcre_free_study
#define pcret_version						pcre_version
#define pcret_pattern_to_host_byte_order	pcre_pattern_to_host_byte_order
#define pcret_jit_stack_alloc				pcre_jit_stack_alloc
#define pcret_jit_stack_free				pcre_jit_stack_free
#define pcret_assign_jit_stack				pcre_assign_jit_stack

#endif