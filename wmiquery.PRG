**************************************************************
* Marco Plaza, 2018,2021
* @nfoxDev
* https://github.com/nftools/wmiquery
**************************************************************
*
* WMI Query Tool : returns an object with item count & items array from wmiquery
*
* simple usage: wmiQuery( wmiClass [where <filter condition>] [, wmiNameSpace] )
*
* ( wmiNameSpace defaults to "CIMV2" )
*
* Returns: Object as #oWmiResult
*
*#oWmiResult:
*      .count = "i"
*      .items[]
*         -item = "v"
*
* test:
* save this program as wmiquery.prg
* do testme in wmiquery 
* check oInfo using debugger / intellisense
*
*****************************************************************
lParameters wmiClass,wmiNameSpace

Local emessage
Local objwmiservice
Local oquery
Local owmi

wmiClass	 = Evl(m.wmiClass,'')
wmiNameSpace = Evl(m.wmiNameSpace,'CIMV2')

emessage = ''


Try
	objwmiservice = Getobject("winmgmts://./root/"+m.wmiNameSpace)
	oquery	= objwmiservice.execquery( 'SELECT * FROM '+m.wmiClass,,48)
	owmi		= processobject( m.oquery )
Catch To oerr
	emessage = m.oerr.Message
Endtry

If !Empty(m.emessage)
	Error ' Invalid WMI Class or NameSpace '
	Return .Null.
Else
	Return m.owmi
Endif

*-------------------------------------------------
Procedure processobject( oquery )
*-------------------------------------------------
Local owmi
Local nitem

owmi = Createobject('empty')
AddProperty(owmi,'items(1)',.Null.)
nitem = 0

Try

	For Each oitem In m.oquery

		nitem = m.nitem + 1
		Dimension owmi.items(m.nitem)
		owmi.items(m.nitem) = Createobject('empty')
		setproperties( m.oitem, owmi.items(m.nitem) )

	Endfor

Catch

Endtry

AddProperty(owmi,'count',m.nitem)

Return m.owmi

*--------------------------------------------------------
Procedure setproperties( oitem , otarget  )
*--------------------------------------------------------

Local thisproperty
Local thisarray
Local nitem
Local thisitem
Local newName

For Each property In m.oitem.properties_
	Try
		Do Case
		Case Vartype( m.property.Value ) = 'O'
			thisproperty = Createobject('empty')
			setproperties(m.property.Value, m.thisproperty )
			AddProperty( otarget ,m.property.Name,m.thisproperty)

		Case m.property.isarray

			AddProperty( otarget ,property.Name+'(1)',.Null.)
			thisarray = 'otarget.'+m.property.Name

			nitem = 0

			If !Isnull(m.property.Value)

				For Each Item In m.property.Value

					nitem = m.nitem+1
					Dimension &thisarray(m.nitem)

					If Vartype( m.item) = 'O'
						thisitem = Createobject('empty')
						setproperties( m.item, m.thisitem )
						&thisarray(m.nitem) = m.thisitem
					Else
						&thisarray(m.nitem) = m.item
					Endif

				Endfor

			Endif

		Otherwise
			Try
				AddProperty( otarget ,m.property.Name,m.property.Value)
			Catch
				newName =  property.Name+'_'
				AddProperty( otarget ,m.newName,m.property.Value)
			Endtry

		Endcase

	Catch To oerr
		Messagebox( Textmerge('<<oerr.lineno>> <<oerr.message>> '),0)
		Exit
	Endtry
Endfor

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



*---------------------------------------------------------------------------------
Procedure testme_
* note:
* this code uses underscore ( _.prg ) 
* as a replacement for addproperty()
* available at https://raw.githubusercontent.com/nftools/underscore/master/_.prg
*---------------------------------------------------------------------------------
Public oinfo

oinfo = Create('empty')

Wait 'Running WMI Query....please wait.. ' Window Nowait At Wrows()/2,Wcols()/2


With _( m.oinfo )
	.monitors  =  wmiquery('Win32_PNPEntity where service = "monitor"')
	.diskdrive =  wmiquery('Win32_diskDrive')
	.startup   =  wmiquery('Win32_startupCommand')
	.BaseBoard =  wmiquery('Win32_baseBoard')
	.netAdaptersConfig = wmiquery('Win32_NetworkAdapterConfiguration')
Endwith


Messagebox( 'Please explore "oInfo" in debugger watch window or command line ',0)


