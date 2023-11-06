# winrt_toast_notification

## Note

1. This only supports Windows

## This sample project demos

1. Use crate `windows` to get user's toast notification list in action center
   * The outlook Toast Notification uses template as below.

        ```xml
        <toast launch="O00000000C3B53827EACC97428D8882268332D3CB07005BC0F9CBBCB51741A90F2BB0ED8913D800000000010C00005699A6165AB5CD449A587E69451E84D100050770919E0000">
            <visual>
                <binding template="ToastImageAndText04">
                    <image id="1" src="file:///C:\Users\User~1.COM\AppData\Local\Temp\Olktmp61.png" alt="Placeholder image" placement="AppLogoOverride" hint-crop="circle"/>
                    <text id="1">Sender Name</text>
                    <text id="2">Mail subject</text>
                    <text id="3">Mail Content Summary</text>
                </binding>
            </visual>
            <actions>
                <action activationType="background" arguments="D00000000C3B53827EACC97428D8882268332D3CB07005BC0F9CBBCB51741A90F2BB0ED8913D800000000010C00005699A6165AB5CD449A587E69451E84D100050770919E0000" content="Delete" imageUri="file:///C:\Users\User~1.COM\AppData\Local\Temp\ToasttmpD.png"/>
                <action activationType="background" arguments="F00000000C3B53827EACC97428D8882268332D3CB07005BC0F9CBBCB51741A90F2BB0ED8913D800000000010C00005699A6165AB5CD449A587E69451E84D100050770919E0000" content="Flag" imageUri="file:///C:\Users\User~1.COM\AppData\Local\Temp\ToasttmpF.png"/>
                <action activationType="background" arguments="X00000000C3B53827EACC97428D8882268332D3CB07005BC0F9CBBCB51741A90F2BB0ED8913D800000000010C00005699A6165AB5CD449A587E69451E84D100050770919E0000" content="Dismiss" imageUri="file:///C:\Users\User~1.COM\AppData\Local\Temp\ToasttmpX.png"/>
            </actions>
            <audio src="ms-winsoundevent:Notification.Mail" silent="false"/>
        </toast>
        ```

2. Use crate `netcorehost` to call C# 6.0 library which is used for get outlook MailItem by entry id. See: `OutlookInterop/OutlookInterop/mapi.cs`

## Useful Links

1. [Windows.UI.Notifications Namespace](https://learn.microsoft.com/en-us/uwp/api/windows.ui.notifications?view=winrt-22621)
2. [Toast schema](https://learn.microsoft.com/en-us/uwp/schemas/tiles/toastschema/schema-root)
3. [netcorehost documentation](https://docs.rs/netcorehost/0.15.1/netcorehost/index.html)
4. [netcorehost example passing-parameters](https://github.com/OpenByteDev/netcorehost/blob/master/examples/passing-parameters/main.rs)