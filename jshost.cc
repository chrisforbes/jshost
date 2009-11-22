#include <cstdio>
#include <windows.h>
#include <activscp.h>
#include <atlbase.h>
#include <comutil.h>

#include <map>
#include <string>
#include "site.h"

extern "C" const GUID __declspec(selectany) 
	CLSID_JScript = { 0xf414c260, 0x6ac0, 0x11cf, 0xb6, 0xd1, 0x00, 0xaa, 0x00, 0xbb, 0xbb, 0x58 };

extern IHostObjects * CreateSite( IActiveScript * scr );

HRESULT JsDebug( void * target, DISPPARAMS * params, VARIANT * result )
{
	if( params->cArgs == 1 && params->rgvarg[0].vt == VT_BSTR) {
		std::string text = (char const *)(_bstr_t)(_variant_t) params->rgvarg[0];
		puts( text.c_str() );
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT JsLoadScript( char const * filename )
{
	return E_INVALIDARG;
}

HRESULT JsInclude( void * target, DISPPARAMS * params, VARIANT * result )
{
	for( int i = 0; i < params->cArgs; i++ )
		if (params->rgvarg[i].vt != VT_BSTR)
			return DISP_E_BADVARTYPE;

	for( int i = 0; i < params->cArgs; i++ )
	{
		HRESULT result = JsLoadScript( (char const *)(_bstr_t)(_variant_t) params->rgvarg[i] );
		if (FAILED(result))
			return result;
	}

	return S_OK;
}

HRESULT JsInvoke( void * target, DISPPARAMS * params, VARIANT * result )
{
	printf( "JsInvoke called, cArgs=%u\n", params->cArgs );
	if (params->rgvarg[0].vt == VT_DISPATCH)
		printf( "JsInvoke arg IDispatch\n" );

	IDispatch * f = params->rgvarg[0].pdispVal;
	_variant_t arg("Hello from JsInvoke!");
	DISPPARAMS dp = { &arg, 0, 1, 0 };
	VARIANT rr;
	f->Invoke( 0, IID_NULL, ::GetUserDefaultLCID(), DISPATCH_METHOD, &dp, &rr, 0, 0 );
	return S_OK;
}

IDispatch * CreateGlobalHostObject() 
{ 
	IDispatchDynamic * d = IDispatchDynamic::Create();
	d->AddMember( L"debug", 0, JsDebug );
	d->AddMember( L"include", 0, JsInclude );
	d->AddMember( L"invoke", 0, JsInvoke );
	return d;
}

int main( int argc, char ** argv )
{
	::CoInitialize( 0 );
	CComPtr<IActiveScript> scriptEngine;
	if (S_OK != scriptEngine.CoCreateInstance( CLSID_JScript ))
		printf( "Failed creating JScript Engine!\n" );

	IHostObjects * hostObjects = CreateSite( scriptEngine );
	scriptEngine->SetScriptSite( hostObjects );

	IDispatch * globals = CreateGlobalHostObject();

	hostObjects->AddObject( L"host", globals );

	CComPtr<IActiveScriptParse> parser;
	if (S_OK == scriptEngine->QueryInterface<IActiveScriptParse>( &parser ))
		printf( "Acquired parser for JS\n" );
	
	parser->InitNew();

	CComPtr<IPersist> parserFile;
	parser.QueryInterface( &parserFile );

	if (parserFile)
		printf( "IPersistFile acquired\n" );

	EXCEPINFO ei;
	if (S_OK == parser->ParseScriptText( 
		L"host.invoke( function(x) { host.debug(x); } ); \n"
		L"host.debug( \"Hello from Javascript!\" ); \n",
		0, 
		0, 0, 0, 0, 0L, 0, &ei ))
		printf( "Source parsed.\n" );

	scriptEngine->SetScriptState( SCRIPTSTATE_CONNECTED );
	scriptEngine->SetScriptSite(0);
}