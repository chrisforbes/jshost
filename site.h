#pragma once

class __declspec(novtable) IHostObjects : public IActiveScriptSite
{
public:
	virtual void AddObject( std::wstring const & w, IUnknown * p ) = 0;
};

class __declspec(novtable) IDispatchDynamic : public IDispatch
{
public:
	typedef HRESULT DispatchFunc( void * target, DISPPARAMS * params, VARIANT * result );
	virtual void AddMember( std::wstring const & name, void * target, DispatchFunc * f ) = 0;
	static IDispatchDynamic * Create();
};