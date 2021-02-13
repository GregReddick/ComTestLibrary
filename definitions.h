

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Mon Jan 18 19:14:07 2038
 */
/* Compiler settings for D:\src.cs\ComTestLibrary\definitions.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __definitions_h__
#define __definitions_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IComTest_FWD_DEFINED__
#define __IComTest_FWD_DEFINED__
typedef interface IComTest IComTest;

#endif 	/* __IComTest_FWD_DEFINED__ */


#ifndef __ComTest_FWD_DEFINED__
#define __ComTest_FWD_DEFINED__

#ifdef __cplusplus
typedef class ComTest ComTest;
#else
typedef struct ComTest ComTest;
#endif /* __cplusplus */

#endif 	/* __ComTest_FWD_DEFINED__ */


/* header files for imported files */
#include "unknwn.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IComTest_INTERFACE_DEFINED__
#define __IComTest_INTERFACE_DEFINED__

/* interface IComTest */
/* [oleautomation][unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IComTest;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1B31B683-F0AA-4E71-8F50-F2D2E5E9E210")
    IComTest : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ComTestMethod( 
            /* [in] */ double radius,
            /* [in] */ BSTR comment,
            /* [retval][out] */ double *ReturnVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IComTestVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IComTest * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IComTest * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IComTest * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IComTest * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IComTest * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IComTest * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IComTest * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ComTestMethod )( 
            IComTest * This,
            /* [in] */ double radius,
            /* [in] */ BSTR comment,
            /* [retval][out] */ double *ReturnVal);
        
        END_INTERFACE
    } IComTestVtbl;

    interface IComTest
    {
        CONST_VTBL struct IComTestVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IComTest_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IComTest_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IComTest_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IComTest_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IComTest_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IComTest_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IComTest_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IComTest_ComTestMethod(This,radius,comment,ReturnVal)	\
    ( (This)->lpVtbl -> ComTestMethod(This,radius,comment,ReturnVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IComTest_INTERFACE_DEFINED__ */



#ifndef __ComTestLibrary_LIBRARY_DEFINED__
#define __ComTestLibrary_LIBRARY_DEFINED__

/* library ComTestLibrary */
/* [helpstring][uuid] */ 


EXTERN_C const IID LIBID_ComTestLibrary;

EXTERN_C const CLSID CLSID_ComTest;

#ifdef __cplusplus

class DECLSPEC_UUID("71AD0B2F-E5D0-4272-A4FD-18F707D5E0D6")
ComTest;
#endif
#endif /* __ComTestLibrary_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


