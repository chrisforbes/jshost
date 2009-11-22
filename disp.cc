#include <cstdio>
#include <windows.h>
#include <activscp.h>
#include <atlbase.h>
#include <comutil.h>

#include <map>
#include <string>
#include "site.h"

#pragma comment( lib, "comsuppw.lib" )

struct DispatchInfo
{
	void * target;
	IDispatchDynamic::DispatchFunc * f;
};

class DispatchImpl : public IDispatchDynamic
{
	long _refs;
	std::map< std::wstring, DISPID > _dispids;
	std::map< DISPID, DispatchInfo > _dispinfos;
	DISPID _nextdisp;

public:
	// IUnknown bits
	STDMETHOD_( ULONG, AddRef )() { return ::InterlockedIncrement( &_refs );}
	STDMETHOD_( ULONG, Release )()
	{
		long r = ::InterlockedDecrement( &_refs );
		if (!r) delete this;
		return r;
	}

	STDMETHOD(QueryInterface)( IID const & iid, void ** out )
	{
		if (iid == __uuidof(IUnknown)) { *out = this; return S_OK; }
		if (iid == __uuidof(IDispatch)) { *out = this; return S_OK; }
		return E_NOINTERFACE;
	}

	// IDispatch bits
	STDMETHOD(GetTypeInfoCount)( unsigned * pctInfo ) 
	{ 
		*pctInfo = 0; 
		return S_OK; 
	}

	STDMETHOD(GetTypeInfo)( unsigned index, LCID lcid, ITypeInfo ** ppTypeInfo )
	{
		if (ppTypeInfo) *ppTypeInfo = 0;
		return TYPE_E_ELEMENTNOTFOUND;
	}

	STDMETHOD(GetIDsOfNames)(IID const & iid, OLECHAR ** rgszNames, 
		unsigned n, LCID lcid, DISPID * rcDispIds )
	{
		bool anyFailures = false;
		while( n-- ) {
			std::map< std::wstring, DISPID >::const_iterator it = _dispids.find( *rgszNames++ );
			*rcDispIds++ = (it == _dispids.end()) ? DISPID_UNKNOWN : it->second;
			if (it == _dispids.end()) anyFailures = true;
		}
		return anyFailures ? DISP_E_UNKNOWNNAME : S_OK;
	}

	STDMETHOD(Invoke)( DISPID member, IID const & iid, LCID lcid, WORD flags, 
		DISPPARAMS * params,
		VARIANT * result,
		EXCEPINFO * ei,
		unsigned * ea )
	{
		std::map< DISPID, DispatchInfo >::const_iterator it 
			= _dispinfos.find( member );

		if (it != _dispinfos.end()) 
		{
			return it->second.f( it->second.target, 
				params, result );
		}

		return DISP_E_MEMBERNOTFOUND;
	}

	// IDispatchDynamic bits
	void AddMember( std::wstring const & name, void * target, DispatchFunc * f )
	{
		DISPID id = _nextdisp++;
		DispatchInfo di = { target, f };
		_dispids[ name ] = id;
		_dispinfos[ id ] = di;
	}

	// DispatchImpl bits
	DispatchImpl() : _refs(1), _nextdisp(0x00010000) {}
	
};

IDispatchDynamic * IDispatchDynamic::Create() { return new DispatchImpl(); }