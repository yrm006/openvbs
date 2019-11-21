#include "jujube.h"









bool JSONObject::from(const wchar_t* c){
    while( *c==L' ' || *c==L'\t' || *c==L'\n' ) ++c;
    if(*c == L'{'){ m_map.clear(); ++c; }

    _variant_t* append = nullptr;

    while(*c){
        if(*c == L':'){
            ++c;
        }else
        if(*c == L','){
            ++c;
        }else
        if(*c == L'"'){
            ++c;
            std::wstring s;
            while(!( *c == L'"' )){
                if(!*c){
                    return false;
                }else
                if(*c == L'\\'){
                    ++c;
                    if(!*c){
                        return false;
                    }else
                    if(*c == L'b'){
                        s += L'\b';
                        ++c;
                    }else
                    if(*c == L'f'){
                        s += L'\f';
                        ++c;
                    }else
                    if(*c == L'n'){
                        s += L'\n';
                        ++c;
                    }else
                    if(*c == L'r'){
                        s += L'\r';
                        ++c;
                    }else
                    if(*c == L't'){
                        s += L'\t';
                        ++c;
                    }else
                    if(*c == L'u'){
                        ++c;
                        wchar_t* pend;
                        s += (wchar_t)wcstoll(c, &pend, 10);
                        c = pend;
                    }else{
                        s += *c;
                        ++c;
                    }
                }else{
                    s += *c;
                    ++c;
                }
            }
            ++c;

            if(append){
                append->vt = VT_BSTR;
                append->bstrVal = SysAllocString(s.c_str());
                append = nullptr;
            }else{
                append = &m_map[s];
            }
        }else
        if( append && (*c == L'+' || *c == L'-' || (L'0'<=*c && *c<=L'9')) ){
            wchar_t* pend;
            long long ll = wcstoll(c, &pend, 10);
            if(*pend == L'.'){
                append->vt = VT_R8;
                append->dblVal = wcstod(c, &pend);
            }else{
                append->vt = VT_I8;
                append->llVal = ll;
            }
            append = nullptr;
            c = pend;
        }else
        if( append && (wcsncmp(c, L"true", 4) == 0) ){
            append->vt = VT_BOOL;
            append->boolVal = VARIANT_TRUE;
            append = nullptr;
            c += 4;
        }else
        if( append && (wcsncmp(c, L"false", 5) == 0) ){
            append->vt = VT_BOOL;
            append->boolVal = VARIANT_FALSE;
            append = nullptr;
            c += 5;
        }else
        if( append && (wcsncmp(c, L"null", 4) == 0) ){
            append->vt = VT_NULL;
            append = nullptr;
            c += 4;
        }else
        if( append && (*c == L'{') ){
            const wchar_t* c0 = c++;
            int nest = 1;
            while(nest){
                if(*c==L'{') ++nest;
                else
                if(*c==L'}') --nest;

                ++c;
            }

            append->vt = VT_DISPATCH;
            append->pdispVal = new JSONObject( std::wstring(c0+1, c-c0-2).c_str() );
            append = nullptr;
        }else
        if( append && (*c == L'[') ){
            const wchar_t* c0 = c++;
            int nest = 1;
            while(nest){
                if(*c==L'[') ++nest;
                else
                if(*c==L']') --nest;

                ++c;
            }

            append->vt = VT_DISPATCH;
            append->pdispVal = new JSONArray( std::wstring(c0+1, c-c0-2).c_str() );
            append = nullptr;
        }else
        if( *c == L'}' ){
            break;
        }else
        {
            ++c;
        }
    }

    return true;
}

bool JSONArray::from(const wchar_t* c){
    while( *c==L' ' || *c==L'\t' || *c==L'\n' ) ++c;
    if(*c == L'['){ m_vec.clear(); ++c; }

    while(*c){
        if(*c == L','){
            ++c;
        }else
        if(*c == L'"'){
            ++c;
            std::wstring s;
            while(!( *c == L'"' )){
                if(!*c){
                    return false;
                }else
                if(*c == L'\\'){
                    ++c;
                    if(!*c){
                        return false;
                    }else
                    if(*c == L'b'){
                        s += L'\b';
                        ++c;
                    }else
                    if(*c == L'f'){
                        s += L'\f';
                        ++c;
                    }else
                    if(*c == L'n'){
                        s += L'\n';
                        ++c;
                    }else
                    if(*c == L'r'){
                        s += L'\r';
                        ++c;
                    }else
                    if(*c == L't'){
                        s += L'\t';
                        ++c;
                    }else
                    if(*c == L'u'){
                        ++c;
                        wchar_t* pend;
                        s += (wchar_t)wcstoll(c, &pend, 10);
                        c = pend;
                    }else{
                        s += *c;
                        ++c;
                    }
                }else{
                    s += *c;
                    ++c;
                }
            }
            ++c;

            m_vec.push_back( _variant_t() );
            m_vec.back().vt = VT_BSTR;
            m_vec.back().bstrVal = SysAllocString(s.c_str());
        }else
        if( *c == L'+' || *c == L'-' || (L'0'<=*c && *c<=L'9') ){
            wchar_t* pend;
            long long ll = wcstoll(c, &pend, 10);
            if(*pend == L'.'){
                m_vec.push_back( _variant_t() );
                m_vec.back().vt = VT_R8;
                m_vec.back().dblVal = wcstod(c, &pend);
            }else{
                m_vec.push_back( _variant_t() );
                m_vec.back().vt = VT_I8;
                m_vec.back().llVal = ll;
            }
            c = pend;
        }else
        if( wcsncmp(c, L"true", 4) == 0 ){
            m_vec.push_back( _variant_t() );
            m_vec.back().vt = VT_BOOL;
            m_vec.back().boolVal = VARIANT_TRUE;
            c += 4;
        }else
        if( wcsncmp(c, L"false", 5) == 0 ){
            m_vec.push_back( _variant_t() );
            m_vec.back().vt = VT_BOOL;
            m_vec.back().boolVal = VARIANT_FALSE;
            c += 5;
        }else
        if( wcsncmp(c, L"null", 4) == 0 ){
            m_vec.push_back( _variant_t() );
            m_vec.back().vt = VT_NULL;
            c += 4;
        }else
        if( *c == L'{' ){
            const wchar_t* c0 = c++;
            int nest = 1;
            while(nest){
                if(*c==L'{') ++nest;
                else
                if(*c==L'}') --nest;

                ++c;
            }

            m_vec.push_back( _variant_t() );
            m_vec.back().vt = VT_DISPATCH;
            m_vec.back().pdispVal = new JSONObject( std::wstring(c0+1, c-c0-2).c_str() );
        }else
        if( *c == L'[' ){
            const wchar_t* c0 = c++;
            int nest = 1;
            while(nest){
                if(*c==L'[') ++nest;
                else
                if(*c==L']') --nest;

                ++c;
            }

            m_vec.push_back( _variant_t() );
            m_vec.back().vt = VT_DISPATCH;
            m_vec.back().pdispVal = new JSONArray( std::wstring(c0+1, c-c0-2).c_str() );
        }else
        if( *c == L']' ){
            break;
        }else
        {
            ++c;
        }
    }

    return true;
}









void* CProgram::map_word(const istring& s){
    return CProcessor::map_word(s);
}

const std::map<istring, CProcessor::word_m> CProcessor::s_words = {
    {L"if",            &CProcessor::word_if},
    {L"then",          &CProcessor::word_colon},
    {L"else",          &CProcessor::word_else},
    {L"elseif",        &CProcessor::word_elseif},
    {L"end if",        &CProcessor::word_endif},
    {L"select case",   &CProcessor::word_selectcase},
    {L"case",          &CProcessor::word_case},
    {L"case else",     &CProcessor::word_caseelse},
    {L"end select",    &CProcessor::word_endselect},
    {L"do while",      &CProcessor::word_dowhile},
    {L"while",         &CProcessor::word_dowhile},
    {L"do until",      &CProcessor::word_dountil},
    {L"loop",          &CProcessor::word_loop},
    {L"wend",          &CProcessor::word_loop},
    {L"for each",      &CProcessor::word_foreach},
    {L"in",            &CProcessor::word_in},
    {L"for",           &CProcessor::word_for},
    {L"to",            &CProcessor::word_to},
    {L"step",          &CProcessor::word_step},
    {L"next",          &CProcessor::word_next},
    {L"goto",          &CProcessor::word_goto},
    {L"return",        &CProcessor::word_return},
    {L"exit function", &CProcessor::word_exitfunction},
    {L"exit sub",      &CProcessor::word_exitfunction},
    {L"exit do",       &CProcessor::word_exitdo},
    {L"exit for",      &CProcessor::word_exitfor},
    {L"with",          &CProcessor::word_with},
    {L"end with",      &CProcessor::word_endwith},
    {L"=",             &CProcessor::word_equal},
    {L":=",            &CProcessor::word_coleq},
    {L"+",             &CProcessor::word_plus},
    {L"-",             &CProcessor::word_minus},
    {L"^",             &CProcessor::word_hat},
    {L"*",             &CProcessor::word_aster},
    {L"/",             &CProcessor::word_slash},
    {L"mod",           &CProcessor::word_mod},
    {L"&",             &CProcessor::word_amp},
    {L"++",            &CProcessor::word_inc},
    {L"--",            &CProcessor::word_dec},
    {L"<",             &CProcessor::word_lthan},
    {L">",             &CProcessor::word_gthan},
    {L"<=",            &CProcessor::word_lethan},
    {L">=",            &CProcessor::word_gethan},
    {L"<>",            &CProcessor::word_neq},
    {L"==",            &CProcessor::word_eqeq},
    {L"===",           &CProcessor::word_eqeqeq},
    {L"!==",           &CProcessor::word_exeqeq},
    {L"is",            &CProcessor::word_is},
    {L"not",           &CProcessor::word_not},
    {L"and",           &CProcessor::word_and},
    {L"or",            &CProcessor::word_or},
    {L"xor",           &CProcessor::word_xor},
    {L"\n",            &CProcessor::word_colon},
    {L":",             &CProcessor::word_colon},
    {L".",             &CProcessor::word_dot},
    {L"?.",            &CProcessor::word_quesdot},
    {L"??",            &CProcessor::word_quesques},
    {L",",             &CProcessor::word_comma},
    {L"(",             &CProcessor::word_parenL},
    {L")",             &CProcessor::word_parenR},
    {L"",              &CProcessor::word_exit},
    {L"getref",               &CProcessor::word_getref},
    {L"me",                   &CProcessor::word_me},
    {L"err",                  &CProcessor::word_err},
    {L"new",                  &CProcessor::word_new},
    {L"redim",                &CProcessor::word_redim},
    {L"redim preserve",       &CProcessor::word_redimpreserve},
    {L"on error resume next", &CProcessor::word_onerrorresumenext},
    {L"on error goto catch",  &CProcessor::word_onerrorgotocatch},
    {L"on error catch",       &CProcessor::word_onerrorcatch},
    {L"on error goto 0",      &CProcessor::word_onerrorgoto0},
    {L"@literal",      &CProcessor::word_literal},
    {L"@json_obj",     &CProcessor::word_jsonobj},
    {L"@json_ary",     &CProcessor::word_jsonary},
    {L"@dim",          &CProcessor::word_dim},
    {L"@pdim",         &CProcessor::word_pdim},
    {L"@gdim",         &CProcessor::word_gdim},
    {L"@cdim",         &CProcessor::word_cdim},
    {L"@closure",      &CProcessor::word_closure},
    {L"@extdisp",      &CProcessor::word_extdisp},
    {L"@noop",         &CProcessor::word_noop},
    {L"@func",         &CProcessor::word_func},
    {L"@class",        &CProcessor::word_class},
    {L"@ext",          &CProcessor::word_ext},
    {L"@env",          &CProcessor::word_env},
    {L"@",             &CProcessor::word_},
    {L"_DUMP_",        &CProcessor::word_DUMP},
};

void* CProcessor::map_word(const istring& s){
    auto found = CProcessor::s_words.find(s);
    return (found != CProcessor::s_words.end()) ? (void*)&found->second : (void*)&CProcessor::s_words.find(L"@")->second;
}

Null CProcessor::s_oNull;

const CProcessor::inst_t CProcessor::s_insts[] = {
    &CProcessor::op_copy,
    &CProcessor::op_plus1,
    &CProcessor::op_plus,
    &CProcessor::op_minus1,
    &CProcessor::op_minus,
    &CProcessor::op_hat,
    &CProcessor::op_aster,
    &CProcessor::op_slash,
    &CProcessor::op_mod,
    &CProcessor::op_amp,
    &CProcessor::op_incR,
    &CProcessor::op_incL,
    &CProcessor::op_decR,
    &CProcessor::op_decL,
    &CProcessor::op_lthan,
    &CProcessor::op_gthan,
    &CProcessor::op_lethan,
    &CProcessor::op_gethan,
    &CProcessor::op_neq,
    &CProcessor::op_equal,
    &CProcessor::op_triequal,
    &CProcessor::op_trineq,
    &CProcessor::op_is,
    &CProcessor::op_not,
    &CProcessor::op_and,
    &CProcessor::op_or,
    &CProcessor::op_xor,
    &CProcessor::op_step,
    &CProcessor::op_parenL,
    &CProcessor::op_invoke,
    &CProcessor::op_invoke_put,
    &CProcessor::op_call,
    &CProcessor::op_array,
    &CProcessor::op_array_put,
    &CProcessor::op_getref,
    &CProcessor::op_exit,
    &CProcessor::stmt_if,
    &CProcessor::stmt_selectcase,
    &CProcessor::stmt_case,
    &CProcessor::stmt_caseelse,
    &CProcessor::stmt_do,
    &CProcessor::stmt_foreach,
    &CProcessor::stmt_in,
    &CProcessor::stmt_for,
    &CProcessor::stmt_next,
    &CProcessor::stmt_goto,
    &CProcessor::stmt_return,
    &CProcessor::stmt_with,
    &CProcessor::stmt_new,
    &CProcessor::stmt_redim,
    &CProcessor::stmt_redimpreserve,
};


