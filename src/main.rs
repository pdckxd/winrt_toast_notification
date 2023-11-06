use std::ffi::OsString;

use chrono::{DateTime, Utc};
use chrono::{Local, TimeZone};
use netcorehost::nethost;
use netcorehost::pdcstr;
use netcorehost::pdcstring::PdCString;
// use anyhow::{anyhow, Context, Result};
use windows::{core::Error, UI::Notifications::ToastNotificationManager};

#[allow(dead_code)]
#[derive(Debug, Clone)]
struct OutlookNotification {
    mail_entry_id: String,
    sender_name: String,
    mail_subject: String,
    mail_partial_content: String,
    sent_on: String,
}

impl OutlookNotification {
    pub fn new(
        mail_entry_id: &str,
        sender_name: &str,
        mail_subject: &str,
        mail_partial_content: &str,
        sent_on: &str,
    ) -> Self {
        Self {
            mail_entry_id: mail_entry_id.to_string(),
            sender_name: sender_name.to_string(),
            mail_subject: mail_subject.to_string(),
            mail_partial_content: mail_partial_content.to_string(),
            sent_on: sent_on.to_string(),
        }
    }
}

#[allow(dead_code)]
enum UserNotification {
    Slack,
    Outlook(Vec<OutlookNotification>),
}

fn main() -> Result<(), Error> {
    let hostfxr = nethost::load_hostfxr().unwrap();
    let context = hostfxr
        .initialize_for_runtime_config(pdcstr!(r"OutlookInterop\OutlookInterop\bin\Debug\net6.0\OutlookInterop.runtimeconfig.json"))
        .unwrap();
    // let dll_path = pdcstr!(r"OutlookInterop\OutlookInterop\bin\Debug\net6.0\OutlookInterop.runtimeconfig.json");
    // println!("{}", dll_path);
    let fn_loader = context
        .get_delegate_loader_for_assembly(PdCString::from_os_str(OsString::from(r"OutlookInterop\OutlookInterop\bin\Debug\net6.0\OutlookInterop.dll")).unwrap())
        //     "c:\\OutlookInterop.dll"
        // ))
        .unwrap();
    let fn_get_mail_receive_date = fn_loader
        .get_function_with_unmanaged_callers_only::<fn(text_ptr: *const u8, text_length: i32) -> i64>(
            pdcstr!("OutlookInterop.Mapi, OutlookInterop"),
            pdcstr!("GetMailReceiveDate"),
        )
        .unwrap();
    // let mail_entry_id = "00000000C3B53827EACC97428D8882268332D3CB07005BC0F9CBBCB51741A90F2BB0ED8913D800000000010C00005699A6165AB5CD449A587E69451E84D10005077092EA0000";
    // let timestamp = fn_get_mail_receive_date(mail_entry_id.as_ptr(), mail_entry_id.len() as i32); // prints "Hello from C#!"
    //                                                                                         // println!("{}", timestamp);
    // let received_datetime = Utc.timestamp_opt(timestamp, 0).unwrap();
    // let local_received_datetime: DateTime<Local> = DateTime::from(received_datetime);

    let history = ToastNotificationManager::History()?;
    let history_list = history.GetHistoryWithId(&"Microsoft.Office.OUTLOOK.EXE.15".into())?;
    println!("{}", history_list.Size()?);

    let mut outlook_notif_list: Vec<OutlookNotification> = Vec::new();

    for item in &history_list {
        if let Ok(content) = &item.Content() {
            let mail_entry_id = content
                // println!("{:?}", mail_entry_id)
                .FirstChild()?
                .Attributes()?
                .GetNamedItem(&"launch".into())?
                .InnerText()?; //.with_context(||format!("getting mail_entry_id"))?;
            let mail_title_node = content.SelectSingleNode(&"//binding/text[2]".into())?; //.with_context(||format!("getting mail_title_node"))?;
            let mail_sender_node = content.SelectSingleNode(&"//binding/text[1]".into())?;
            let mail_content_node = content.SelectSingleNode(&"//binding/text[3]".into())?; //.with_context(||format!("getting mail_title_node"))?;
                                                                                            //.with_context(|| format!("getting mail sender node")).with_context(||format!("gettting mail_sender_node"))?;
            let mail_notfi_title = &mail_title_node.InnerText()?; //.with_context(||format!("getting_mail_notfi_title"))?;
            let mail_notfi_sender = &mail_sender_node.InnerText()?; //.with_context(||format!("getting mail_notfi_sender"))?;
            let mail_notfi_content = &mail_content_node.InnerText()?; //.with_context(||format!("getting mail_notfi_sender"))?;

            // println!("{}", &mail_notfi_title);

            let mail_entry_id = mail_entry_id.to_string_lossy();
            let mail_entry_id = mail_entry_id.as_str().trim_start_matches('O');
            let timestamp = fn_get_mail_receive_date(mail_entry_id.as_ptr(), mail_entry_id.len() as i32); // prints "Hello from C#!"
            let result = match timestamp {
                0 => "Failed to get mail entry",
                _ => "Sucessfully",
            };
            println!("{mail_notfi_title} => {:?} -- {result}", mail_entry_id);
                                                                                                    // println!("{}", timestamp);
            let received_datetime = Utc.timestamp_opt(timestamp, 0).unwrap();
            let local_received_datetime: DateTime<Local> = DateTime::from(received_datetime);

            outlook_notif_list.push(OutlookNotification::new(
                mail_entry_id, // strip 'O' at the beginning
                mail_notfi_sender.to_string_lossy().as_str(),
                mail_notfi_title.to_string_lossy().as_str(),
                mail_notfi_content.to_string_lossy().as_str(), // &Utc::now(),
                local_received_datetime.to_string().as_str(),
            ));
        }
    }

    println!("{:#?}", outlook_notif_list);

    // let outlook_notif_list = &history_list.into_iter().map(|item| {
    //     let mail_entry_id = &item.Content()?.FirstChild()?.Attributes()?.GetNamedItem(&"launch".into())?.InnerText()?;
    // }).collect::<Result<>();

    // let _tag = &history_list.First()?.nth(0).ok_or_else(|| {anyhow!("Failed to get fist element in history list!")})?.Tag()?;
    // // println!("{}", tag.to_string_lossy());
    // let notif_content =&history_list.First()?.nth(0).ok_or(anyhow!("Failed to get first element in history list!"))?.Content()?;
    // // let xml_tpl_str = &notif_content.GetXml().unwrap().to_string_lossy();
    // // println!("{}", xml_tpl_str);

    // // let mail_entry_id = &notif_content;
    // let mail_entry_id = &notif_content.FirstChild()?.Attributes()?.GetNamedItem(&"launch".into())?.InnerText()?;
    // let mail_title_node = &notif_content.SelectSingleNode(&"//binding/text[2]".into())?;
    // let mail_sender_node = &notif_content.SelectSingleNode(&"//binding/text[1]".into()).with_context(||format!("getting mail sender node"))?;
    // let mail_notfi_title = &mail_title_node.InnerText()?;
    // let mail_notfi_sender = &mail_sender_node.InnerText()?;
    // println!("{}", mail_notfi_title);
    // println!("{}", mail_notfi_sender);
    // println!("{}", mail_entry_id);

    // let assembly_path = pdcstr!(r"OutlookInterop\OutlookInterop\bin\Debug\net6.0\OutlookInterop.dll").to_os_string();
    // if Path::new(&assembly_path).exists() {
    //     println!("Found path");
    // }

    // let fn_get_mail_receive_date = fn_loader
    //     .get_function_with_unmanaged_callers_only::<fn(text_ptr: *const u8, text_length: i32) -> i64>(
    //         pdcstr!("OutlookInterop.Mapi, OutlookInterop"),
    //         pdcstr!("GetMailReceiveDate"),
    //     )
    //     .unwrap();
    // let mail_entry_id = "00000000C3B53827EACC97428D8882268332D3CB07005BC0F9CBBCB51741A90F2BB0ED8913D800000000010C00005699A6165AB5CD449A587E69451E84D10005077092EA0000";
    // let timestamp = fn_get_mail_receive_date(mail_entry_id.as_ptr(), mail_entry_id.len() as i32); // prints "Hello from C#!"
    //                                                                                         // println!("{}", timestamp);
    // let received_datetime = Utc.timestamp_opt(timestamp, 0).unwrap();
    // let local_received_datetime: DateTime<Local> = DateTime::from(received_datetime);
    // println!("{}", local_received_datetime);

    Ok(())
}
