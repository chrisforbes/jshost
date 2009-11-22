#include <cstdio>
#include <windows.h>
#include <activscp.h>
#include <atlbase.h>
#include <map>
#include <string>

#include "site.h"

class Site : public IHostObjects
{
	long _refs;
	IActiveScript * _scr;
	std::map< std::wstring, IUnknown * > _objs;

public:
	STDMETHOD_( ULONG, AddRef )() { return ::InterlockedIncrement( &_refs ); }
	STDMETHOD_( ULONG, Release )()
	{
		long r = ::InterlockedDecrement( &_refs );
		if (!r) {
			printf( "Site Destroyed\n" );
			delete this;
		}

		return r;
	}

	STDMETHOD(QueryInterface)( IID const & iid, void ** out )
	{
		if (iid == __uuidof(IUnknown)) { *out = this; return S_OK; }
		if (iid == __uuidof(IActiveScriptSite)) { *out = this; return S_OK; }

		*out = 0;
		return E_NOINTERFACE;
	}

	STDMETHOD(GetLCID)( LCID * plcid ) { return E_NOTIMPL; }
	
	STDMETHOD(GetItemInfo)( LPCOLESTR pstrName, DWORD dwReturnMask, 
		IUnknown ** ppUnkItem, ITypeInfo ** ppTypeInfo )
	{
		if (ppTypeInfo) *ppTypeInfo = 0;
		if (ppUnkItem) *ppUnkItem = 0;

		if ((dwReturnMask & SCRIPTINFO_ALL_FLAGS) == SCRIPTINFO_IUNKNOWN) {
			std::map< std::wstring, IUnknown * >::const_iterator it;
			for( it = _objs.begin(); it != _objs.end(); it++ )		// todo: sequential lookup in map sucks.
				if (!_wcsicmp( it->first.c_str(), pstrName ))		// (for perf)
				{
					*ppUnkItem = it->second;
					it->second->AddRef();
					return S_OK;
				}
		}
		
		return TYPE_E_ELEMENTNOTFOUND;
	}

	STDMETHOD(GetDocVersionString)(BSTR * versionString) { return E_NOTIMPL; }

	STDMETHOD(OnScriptTerminate)( VARIANT const * pvarResult, EXCEPINFO const * pe )
	{
		printf( "OnScriptTerminate called\n" );
		return S_OK;
	}

	STDMETHOD(OnStateChange)( SCRIPTSTATE ss )
	{
		printf( "Script Engine has entered state %u\n", ss );
		return S_OK;
	}

	STDMETHOD(OnScriptError)( IActiveScriptError * e )
	{
		printf( "Script Error TODO\n" );
		return S_OK;
	}

	STDMETHOD(OnEnterScript)( void )
	{
		printf( "Entered script code\n" );
		return S_OK;
	}

	STDMETHOD(OnLeaveScript)( void )
	{
		printf( "Left script code\n" );
		return S_OK;
	}

	void AddObject( std::wstring const & w, IUnknown * p ) 
	{ 
		_objs[ w ] = p; 
		_scr->AddNamedItem( w.c_str(), SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE );
	}

	explicit Site( IActiveScript * scr ) : _refs(1), _scr(scr) {}
};

IHostObjects * CreateSite( IActiveScript * scr ) { return new Site(scr); }