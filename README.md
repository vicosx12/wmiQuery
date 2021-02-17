# WMIQUERY Function

Allows you to make synchronous Wmi Querys over local PC and get result as a VFP object 
using a simple function call.

     wmiQuery( wmiClass [where <filter condition>] [, wmiNameSpace] )
 
### Parameters

**wmiClass**

Any valid WMI Class ( check MS WMI code creator https://www.microsoft.com/en-us/download/details.aspx?id=8572 for complete list ) 

**wmiNameSpace**

Specify the wmiNameSpace; defaults to "CIMV2"

### Return Value: 

Object. 


	Result object schema: 

    .oWmiResult:
      .count = "i"
      .items[]
        -item = "v"
    
### Sample procedure ( included in wmiquery.prg ): 

    
	*----------------------------------
	Procedure testme
	*----------------------------------
	Public oinfo

	oinfo = Create('empty')

	Wait 'Running WMI Query....please wait.. ' Window Nowait At Wrows()/2,Wcols()/2


	AddProperty( oinfo, "monitors"  , wmiquery('Win32_PNPEntity where service = "monitor"') )
	AddProperty( oinfo, "diskdrive" , wmiquery('Win32_diskDrive') )
	AddProperty( oinfo, "startup" ,   wmiquery('Win32_startupCommand'))
	AddProperty( oinfo, "BaseBoard" , wmiquery('Win32_baseBoard') )
	AddProperty( oinfo, "netAdaptersConfig",  wmiquery('Win32_NetworkAdapterConfiguration') )


	Messagebox( 'Please explore "oInfo" in debugger watch window or command line ',0)
    


![](https://github.com/nftools/wmiQuery/blob/master/wmiquery.jpg)

