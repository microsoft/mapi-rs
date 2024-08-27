use proc_macro::TokenStream;
use quote::{format_ident, quote};
use syn::{
    braced,
    parse::{Parse, ParseStream},
    parse_macro_input,
    punctuated::{Pair, Punctuated},
    token::Comma,
    Abi, Expr, ExprLit, FnArg, ForeignItemFn, Ident, Lit, LitStr, Meta, MetaNameValue, Pat,
    PatType, Result, ReturnType,
};

struct DelayLoadAttr {
    pub name: LitStr,
}

impl Parse for DelayLoadAttr {
    fn parse(input: ParseStream) -> Result<Self> {
        let meta: Meta = input.parse()?;
        match meta {
            Meta::NameValue(MetaNameValue {
                path,
                value:
                    Expr::Lit(ExprLit {
                        lit: Lit::Str(name),
                        ..
                    }),
                ..
            }) if path.get_ident().map(Ident::to_string).as_deref() == Some("name") => {
                Ok(DelayLoadAttr { name: name.clone() })
            }
            _ => Err(input.error(r#"expected #[delay_load(name = "...")]"#)),
        }
    }
}

struct ExternDecl {
    pub abi: LitStr,
    pub ident: Ident,
    pub inputs: Punctuated<FnArg, Comma>,
    pub output: ReturnType,
}

impl Parse for ExternDecl {
    fn parse(input: ParseStream) -> Result<Self> {
        let abi: Abi = input.parse()?;
        let abi = abi
            .name
            .ok_or_else(|| input.error(r#"expected "system" or "cdecl""#))?;

        let content;
        braced!(content in input);
        let foreign_item: ForeignItemFn = content.parse()?;

        Ok(ExternDecl {
            abi,
            ident: foreign_item.sig.ident,
            inputs: foreign_item.sig.inputs,
            output: foreign_item.sig.output,
        })
    }
}

/// Implement a delay load helper for the foreign function declaration in an extern block.
#[proc_macro_attribute]
pub fn delay_load(attr: TokenStream, input: TokenStream) -> TokenStream {
    let attr = parse_macro_input!(attr as DelayLoadAttr);
    let ast = parse_macro_input!(input as ExternDecl);
    impl_delay_load(&attr, &ast)
}

fn no_arg_size(undecorated: &str) -> bool {
    use std::{collections::BTreeSet, sync::OnceLock};

    static NO_ARG_SIZE_MAPI: OnceLock<BTreeSet<&'static str>> = OnceLock::new();
    let no_arg_size_mapi = NO_ARG_SIZE_MAPI.get_or_init(|| {
        BTreeSet::from([
            // "BMAPIAddress",
            // "BMAPIDetails",
            // "BMAPIFindNext",
            // "BMAPIGetAddress",
            // "BMAPIGetReadMail",
            // "BMAPIReadMail",
            // "BMAPIResolveName",
            // "BMAPISaveMail",
            // "BMAPISendMail",
            // "FGetComponentPath",
            "FixMAPI",
            "GetOutlookVersion",
            // "GetTnefStreamCodepage",
            "HrGetOmiProvidersFlags",
            "HrSetOmiProvidersFlagsInvalid",
            // "LAUNCHWIZARD",
            // "MAPIAddress",
            // "MAPIAdminProfiles",
            // "MAPIAllocateBuffer",
            // "MAPIAllocateMore",
            // "MAPIDeleteMail",
            // "MAPIDetails",
            // "MAPIFindNext",
            // "MAPIFreeBuffer",
            // "MAPIInitialize",
            // "MAPILogoff",
            // "MAPILogon",
            // "MAPILogonEx",
            // "MAPIOpenFormMgr",
            // "MAPIOpenLocalFormContainer",
            // "MAPIReadMail",
            // "MAPIResolveName",
            // "MAPISaveMail",
            // "MAPISendDocuments",
            // "MAPISendMail",
            // "MAPISendMailW",
            // "MAPIUninitialize",
            // "OpenStreamOnFile",
            // "OpenTnefStream",
            // "OpenTnefStreamEx",
            // "PRProviderInit",
            // "RTFSync",
            // "ScMAPIXFromCMC",
            // "ScMAPIXFromSMAPI",
            // "WrapCompressedRTFStream",
        ])
    });

    static NO_ARG_SIZE_OLMAPI: OnceLock<BTreeSet<&'static str>> = OnceLock::new();
    let no_arg_size_olmapi = NO_ARG_SIZE_OLMAPI.get_or_init(|| {
        BTreeSet::from([
            "BMAPIAddress",
            "BMAPIDetails",
            "BMAPIFindNext",
            "BMAPIGetAddress",
            "BMAPIGetReadMail",
            "BMAPIReadMail",
            "BMAPIResolveName",
            "BMAPISaveMail",
            "BMAPISendMail",
            "ClosePerformanceData",
            "CollectPerformanceData",
            "CreateMapiInitializationMonitor",
            "CreateObject",
            "DoDeliveryReport",
            "EndBoot",
            "EtwTraceMessage",
            "FGetComponentPath",
            "GetTnefStreamCodepage",
            "HrEnsureProviderResourceDLL",
            "HrGetDefaultStoragePathA",
            "HrGetDefaultStoragePathW",
            "HrGetEDPIdentifierFromStoreEIDOnMapi",
            "HrGetOpenTnefStream",
            "HrGetProviderResourceDLL",
            "HrNotify",
            "LAUNCHWIZARD",
            "MAPIAddress",
            "MAPIAdminProfiles",
            "MAPIAllocateBuffer",
            "MAPIAllocateBufferProv",
            "MAPIAllocateMore",
            "MAPIAllocateMoreProv",
            "MAPICrashRecovery",
            "MAPIDeleteMail",
            "MAPIDetails",
            "MAPIFindNext",
            "MAPIFreeBuffer",
            "MAPIInitialize",
            "MAPILogoff",
            "MAPILogon",
            "MAPILogonEx",
            "MAPIOpenFormMgr",
            "MAPIOpenLocalFormContainer",
            "MAPIReadMail",
            "MAPIResolveName",
            "MAPISaveMail",
            "MAPISendDocuments",
            "MAPISendMail",
            "MAPISendMailW",
            "MAPIUninitialize",
            "MAPIValidateAllocatedBuffer",
            "MSProviderInit",
            "OpenPerformanceData",
            "OpenStreamOnFile",
            "OpenStreamOnFileW",
            "OpenTnefStream",
            "OpenTnefStreamEx",
            "OverrideMAPIResourcePath",
            "PRProviderInit",
            "RPCTRACE",
            "RTFSync",
            "RTFSyncCpid",
            "RopString",
            "RpcTraceReadRegSettings",
            "ScMAPIXFromCMC",
            "ScMAPIXFromSMAPI",
            "Unload",
            "WrapCompressedRTFStream",
            "WrapCompressedRTFStreamEx",
            "fnevString",
            "g_dwRpcThreshold",
        ])
    });

    no_arg_size_mapi.contains(undecorated) || no_arg_size_olmapi.contains(undecorated)
}

fn impl_delay_load(attr: &DelayLoadAttr, ast: &ExternDecl) -> TokenStream {
    let dll = &attr.name.value();
    let abi = &ast.abi;
    let name = &ast.ident;
    let inputs = &ast.inputs;
    let output = &ast.output;

    let mut args_size = quote! { 0 };
    let mut forward_args: Punctuated<Box<Pat>, Comma> = Punctuated::new();
    for pair in inputs.pairs() {
        match pair {
            Pair::Punctuated(FnArg::Typed(PatType { pat, ty, .. }), comma) => {
                forward_args.push_value(pat.clone());
                forward_args.push_punct(*comma);
                args_size = quote! { #args_size + mem::size_of::<#ty>() };
            }
            Pair::End(FnArg::Typed(PatType { pat, ty, .. })) => {
                forward_args.push_value(pat.clone());
                args_size = quote! { #args_size + mem::size_of::<#ty>() };
            }
            _ => panic!("should not have a receiver/self argument"),
        }
    }

    let func_type = format_ident!("PFN{}", name);
    let proc_name = LitStr::new(&format!("{name}"), name.span());

    let undecorated = format!("{name}");
    let build_proc_name = if no_arg_size(undecorated.as_str()) {
        quote! {
            let proc_name = s!(#proc_name);
        }
    } else {
        quote! {
            let mut proc_name: Vec<_> = #proc_name.bytes().collect();
            #[cfg(target_pointer_width = "32")]
            {
                const ARG_SIZE: usize = #args_size;
                proc_name.extend(format!("@{ARG_SIZE}").bytes());
            }
            proc_name.push(0);
            let proc_name = PCSTR::from_raw(proc_name.as_ptr());
        }
    };

    let call_export = if dll.as_str() == "olmapi32" {
        quote! {
            static EXPORT: OnceLock<Option<#func_type>> = OnceLock::new();

            use ::windows::Win32::{Foundation::E_FAIL, System::LibraryLoader::*};

            match (EXPORT.get_or_init(|| {
                unsafe {
                    let module = crate::get_mapi_module();
                    GetProcAddress(module, proc_name).map(|export| unsafe { mem::transmute(export) })
                }
            })) {
                Some(export) => {
                    unsafe {
                        export(#forward_args)
                    }
                },
                None => E_FAIL
            }
        }
    } else {
        let missing_export =
            LitStr::new(&format!("{name} is not exported from {dll}"), name.span());

        quote! {
            static EXPORT: OnceLock<#func_type> = OnceLock::new();

            (EXPORT.get_or_init(|| {
                use ::windows::Win32::System::LibraryLoader::*;

                unsafe {
                    let module = crate::get_mapi_module();
                    mem::transmute(GetProcAddress(module, proc_name).expect(#missing_export))
                }
            }))(#forward_args)
        }
    };

    let gen = quote! {
        unsafe fn #name(#inputs) #output {
            use std::{mem, sync::OnceLock};
            use ::windows_core::*;

            #build_proc_name

            type #func_type = unsafe extern #abi fn(#inputs) #output;

            #call_export
        }
    };

    gen.into()
}
