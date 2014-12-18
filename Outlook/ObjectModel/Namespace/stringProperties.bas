' runVBAFilesInOffice.vbs -excel -ol stringProperties -c main

option explicit

sub main()

    dim ol        as outlook.application
    dim ns        as outlook.namespace
    dim r         as long
    dim connMode  as olExchangeConnectionMode 
    dim connMode_ as string

    set ol   = new outlook.application
    set ns   = ol.getNamespace("MAPI")

    r = 1


    connMode = ns.exchangeConnectionMode
    if     connMode    = olCachedConnectedDrizzle  then
           connMode_   ="olCachedConnectedDrizzle"
    elseif connMode    = olCachedConnectedFull     then
           connMode_   ="olCachedConnectedFull"
    elseif connMode    = olCachedConnectedHeaders  then
           connMode_   ="olCachedConnectedHeaders"
    elseif connMode    = olCachedDisconnected      then
           connMode_   ="olCachedDisconnected"
    elseif connMode    = olCachedOffline           then
           connMode_   ="olCachedOffline"
    elseif connMode    = olDisconnected            then
           connMode_   ="olDisconnected"
    elseif connMode    = olNoExchange              then
           connMode_   ="olNoExchange"
    elseif connMode    = olOffline                 then
           connMode_   ="olOffline"
    elseif connMode    = olOnline                  then
           connMode_   ="olOnline"
    end if


    cells(r,1).value = "currentUser"                  : cells(r,2).value = ns.currentUser                  : r = r+1
    cells(r,1).value = "currentProfileName"           : cells(r,2).value = ns.currentProfileName           : r = r+1
    cells(r,1).value = "defaultStore"                 : cells(r,2).value = ns.defaultStore                 : r = r+1
    cells(r,1).value = "exchangeConnectionMode"       : cells(r,2).value =    connMode_                    : r = r+1
    cells(r,1).value = "exchangeMailboxServerName"    : cells(r,2).value = ns.exchangeMailboxServerName    : r = r+1
    cells(r,1).value = "exchangeMailboxServerVersion" : cells(r,2).value = ns.exchangeMailboxServerVersion : r = r+1
    cells(r,1).value = "offline"                      : cells(r,2).value = ns.offline                      : r = r+1
    cells(r,1).value = "type"                         : cells(r,2).value = ns.type                         : r = r+1

    columns(1).autofit

    activeWorkbook.saved = true
end sub
