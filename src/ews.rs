//! Exchange Web Services (EWS) client for Microsoft Exchange/Office 365
//!
//! Uses SOAP/XML over HTTPS with OAuth2 Bearer tokens. Works with both
//! personal and enterprise Microsoft accounts, including tenants that
//! have blocked Graph API and IMAP.
//!
//! # Configuration
//!
//! ```text
//! MAIL_EWS_<SEGMENT>_USER=user@company.com
//! MAIL_EWS_<SEGMENT>_CLIENT_ID=d3590ed6-52b3-4102-aeff-aad2292ab01c
//! MAIL_EWS_<SEGMENT>_CLIENT_SECRET=none
//! MAIL_EWS_<SEGMENT>_REFRESH_TOKEN=<token>
//! ```

use std::time::Duration;

use crate::errors::{AppError, AppResult};
use crate::oauth2::TokenManager;

/// EWS endpoint
const EWS_URL: &str = "https://outlook.office365.com/EWS/Exchange.asmx";

/// EWS XML namespaces
const SOAP_NS: &str = "http://schemas.xmlsoap.org/soap/envelope/";
const TYPES_NS: &str = "http://schemas.microsoft.com/exchange/services/2006/types";
const MESSAGES_NS: &str = "http://schemas.microsoft.com/exchange/services/2006/messages";

// ─── EWS account config ─────────────────────────────────────────────────────

/// EWS account configuration
#[derive(Debug, Clone)]
pub struct EwsAccountConfig {
    pub account_id: String,
    pub user: String,
}

// ─── Response types ──────────────────────────────────────────────────────────

/// A message from EWS FindItem
#[derive(Debug, Clone, serde::Serialize)]
pub struct EwsMessage {
    pub item_id: String,
    pub change_key: String,
    pub subject: String,
    pub from_name: String,
    pub from_email: String,
    pub date_received: String,
    pub is_read: bool,
}

/// A message body from EWS GetItem
#[derive(Debug, Clone, serde::Serialize)]
pub struct EwsMessageDetail {
    pub item_id: String,
    pub subject: String,
    pub from_name: String,
    pub from_email: String,
    pub to: String,
    pub cc: String,
    pub date_received: String,
    pub body_text: String,
    pub is_read: bool,
    pub has_attachments: bool,
}

// ─── Client ──────────────────────────────────────────────────────────────────

/// Send a SOAP request to EWS with OAuth2 Bearer token.
async fn ews_request(
    token_manager: &TokenManager,
    account_id: &str,
    soap_body: &str,
) -> AppResult<String> {
    let access_token = token_manager.get_access_token(account_id).await?;

    let envelope = format!(
        r#"<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="{SOAP_NS}"
               xmlns:t="{TYPES_NS}"
               xmlns:m="{MESSAGES_NS}">
  <soap:Body>
    {soap_body}
  </soap:Body>
</soap:Envelope>"#
    );

    let client = reqwest::Client::new();
    let response = client
        .post(EWS_URL)
        .header("Content-Type", "text/xml")
        .bearer_auth(&access_token)
        .body(envelope)
        .timeout(Duration::from_secs(30))
        .send()
        .await
        .map_err(|e| AppError::Internal(format!("EWS request failed: {e}")))?;

    if !response.status().is_success() {
        let status = response.status();
        let body = response.text().await.unwrap_or_default();
        if status.as_u16() == 401 || status.as_u16() == 403 {
            return Err(AppError::AuthFailed(format!(
                "EWS authentication failed ({status})"
            )));
        }
        return Err(AppError::Internal(format!(
            "EWS request failed ({status}): {body}"
        )));
    }

    response
        .text()
        .await
        .map_err(|e| AppError::Internal(format!("EWS response read failed: {e}")))
}

// ─── Operations ──────────────────────────────────────────────────────────────

/// List messages in a folder (default: inbox)
pub async fn find_items(
    token_manager: &TokenManager,
    account_id: &str,
    folder: &str,
    max_items: usize,
    offset: usize,
) -> AppResult<Vec<EwsMessage>> {
    let folder_id = match folder.to_ascii_lowercase().as_str() {
        "inbox" => "inbox",
        "sent" | "sentitems" | "sent items" => "sentitems",
        "drafts" => "drafts",
        "deleted" | "deleteditems" => "deleteditems",
        "junk" | "junkemail" => "junkemail",
        _ => folder,
    };

    let is_distinguished = matches!(
        folder_id,
        "inbox" | "sentitems" | "drafts" | "deleteditems" | "junkemail"
    );
    let folder_xml = if is_distinguished {
        format!(r#"<t:DistinguishedFolderId Id="{folder_id}"/>"#)
    } else {
        format!(r#"<t:FolderId Id="{folder_id}"/>"#)
    };

    let soap = format!(
        r#"<m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
          <t:FieldURI FieldURI="message:From"/>
          <t:FieldURI FieldURI="message:IsRead"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="{max_items}" Offset="{offset}" BasePoint="Beginning"/>
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        {folder_xml}
      </m:ParentFolderIds>
    </m:FindItem>"#
    );

    let xml = ews_request(token_manager, account_id, &soap).await?;
    parse_find_items_response(&xml)
}

/// Get full message details
pub async fn get_item(
    token_manager: &TokenManager,
    account_id: &str,
    item_id: &str,
) -> AppResult<EwsMessageDetail> {
    let soap = format!(
        r#"<m:GetItem>
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Body"/>
          <t:FieldURI FieldURI="item:HasAttachments"/>
          <t:FieldURI FieldURI="message:ToRecipients"/>
          <t:FieldURI FieldURI="message:CcRecipients"/>
        </t:AdditionalProperties>
        <t:BodyType>Text</t:BodyType>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="{item_id}"/>
      </m:ItemIds>
    </m:GetItem>"#
    );

    let xml = ews_request(token_manager, account_id, &soap).await?;
    parse_get_item_response(&xml)
}

/// Send an email via EWS CreateItem
pub async fn send_email(
    token_manager: &TokenManager,
    account_id: &str,
    to: &[String],
    cc: &[String],
    subject: &str,
    body: &str,
    body_type: &str,
) -> AppResult<()> {
    let to_xml: String = to
        .iter()
        .map(|addr| {
            format!(
                r#"<t:Mailbox><t:EmailAddress>{addr}</t:EmailAddress></t:Mailbox>"#
            )
        })
        .collect();

    let cc_xml: String = cc
        .iter()
        .map(|addr| {
            format!(
                r#"<t:Mailbox><t:EmailAddress>{addr}</t:EmailAddress></t:Mailbox>"#
            )
        })
        .collect();

    let cc_section = if cc.is_empty() {
        String::new()
    } else {
        format!("<t:CcRecipients>{cc_xml}</t:CcRecipients>")
    };

    // Escape XML special characters in body
    let body_escaped = body
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;");
    let subject_escaped = subject
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;");

    let soap = format!(
        r#"<m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:SavedItemFolderId>
        <t:DistinguishedFolderId Id="sentitems"/>
      </m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:Subject>{subject_escaped}</t:Subject>
          <t:Body BodyType="{body_type}">{body_escaped}</t:Body>
          <t:ToRecipients>{to_xml}</t:ToRecipients>
          {cc_section}
        </t:Message>
      </m:Items>
    </m:CreateItem>"#
    );

    let xml = ews_request(token_manager, account_id, &soap).await?;

    // Check for errors in response
    if xml.contains("ResponseClass=\"Error\"") {
        let msg = extract_xml_text(&xml, "MessageText").unwrap_or_default();
        return Err(AppError::Internal(format!("EWS send failed: {msg}")));
    }

    Ok(())
}

// ─── XML Parsing helpers ─────────────────────────────────────────────────────

fn extract_xml_text<'a>(xml: &'a str, tag: &str) -> Option<&'a str> {
    // Simple tag extraction without full XML parser
    let open = format!("<t:{tag}>");
    let close = format!("</t:{tag}>");
    let alt_open = format!("<m:{tag}>");
    let alt_close = format!("</m:{tag}>");

    if let Some(start) = xml.find(&open) {
        let content_start = start + open.len();
        if let Some(end) = xml[content_start..].find(&close) {
            return Some(&xml[content_start..content_start + end]);
        }
    }
    if let Some(start) = xml.find(&alt_open) {
        let content_start = start + alt_open.len();
        if let Some(end) = xml[content_start..].find(&alt_close) {
            return Some(&xml[content_start..content_start + end]);
        }
    }
    None
}

fn extract_attr<'a>(xml: &'a str, tag: &str, attr: &str) -> Option<&'a str> {
    let pattern = format!("<t:{tag} ");
    if let Some(start) = xml.find(&pattern) {
        let attr_pattern = format!("{attr}=\"");
        if let Some(attr_start) = xml[start..].find(&attr_pattern) {
            let val_start = start + attr_start + attr_pattern.len();
            if let Some(end) = xml[val_start..].find('"') {
                return Some(&xml[val_start..val_start + end]);
            }
        }
    }
    None
}

fn parse_find_items_response(xml: &str) -> AppResult<Vec<EwsMessage>> {
    let mut messages = Vec::new();

    // Split by Message tags
    let mut pos = 0;
    while let Some(start) = xml[pos..].find("<t:Message>") {
        let abs_start = pos + start;
        if let Some(end) = xml[abs_start..].find("</t:Message>") {
            let msg_xml = &xml[abs_start..abs_start + end + "</t:Message>".len()];

            let item_id = extract_attr(msg_xml, "ItemId", "Id")
                .unwrap_or_default()
                .to_owned();
            let change_key = extract_attr(msg_xml, "ItemId", "ChangeKey")
                .unwrap_or_default()
                .to_owned();
            let subject = extract_xml_text(msg_xml, "Subject")
                .unwrap_or_default()
                .to_owned();
            let date = extract_xml_text(msg_xml, "DateTimeReceived")
                .unwrap_or_default()
                .to_owned();
            let is_read = extract_xml_text(msg_xml, "IsRead")
                .map(|v| v == "true")
                .unwrap_or(false);

            // Extract from name and email
            let from_name = extract_xml_text(msg_xml, "Name")
                .unwrap_or_default()
                .to_owned();
            let from_email = extract_xml_text(msg_xml, "EmailAddress")
                .unwrap_or_default()
                .to_owned();

            messages.push(EwsMessage {
                item_id,
                change_key,
                subject,
                from_name,
                from_email,
                date_received: date,
                is_read,
            });

            pos = abs_start + end + "</t:Message>".len();
        } else {
            break;
        }
    }

    Ok(messages)
}

fn parse_get_item_response(xml: &str) -> AppResult<EwsMessageDetail> {
    if xml.contains("ResponseClass=\"Error\"") {
        let msg = extract_xml_text(xml, "MessageText").unwrap_or("unknown error");
        return Err(AppError::Internal(format!("EWS GetItem failed: {msg}")));
    }

    let item_id = extract_attr(xml, "ItemId", "Id")
        .unwrap_or_default()
        .to_owned();
    let subject = extract_xml_text(xml, "Subject")
        .unwrap_or_default()
        .to_owned();
    let date = extract_xml_text(xml, "DateTimeReceived")
        .unwrap_or_default()
        .to_owned();
    let is_read = extract_xml_text(xml, "IsRead")
        .map(|v| v == "true")
        .unwrap_or(false);
    let has_attachments = extract_xml_text(xml, "HasAttachments")
        .map(|v| v == "true")
        .unwrap_or(false);

    // Extract body
    let body_text = if let Some(start) = xml.find("<t:Body") {
        if let Some(content_start) = xml[start..].find('>') {
            let abs_start = start + content_start + 1;
            if let Some(end) = xml[abs_start..].find("</t:Body>") {
                xml[abs_start..abs_start + end].to_owned()
            } else {
                String::new()
            }
        } else {
            String::new()
        }
    } else {
        String::new()
    };

    // Extract From
    let from_name = extract_xml_text(xml, "Name")
        .unwrap_or_default()
        .to_owned();
    let from_email = extract_xml_text(xml, "EmailAddress")
        .unwrap_or_default()
        .to_owned();

    // Extract To and CC (simplified)
    let to = extract_recipients(xml, "ToRecipients");
    let cc = extract_recipients(xml, "CcRecipients");

    Ok(EwsMessageDetail {
        item_id,
        subject,
        from_name,
        from_email,
        to,
        cc,
        date_received: date,
        body_text,
        is_read,
        has_attachments,
    })
}

fn extract_recipients(xml: &str, tag: &str) -> String {
    let open = format!("<t:{tag}>");
    let close = format!("</t:{tag}>");
    if let Some(start) = xml.find(&open) {
        let content_start = start + open.len();
        if let Some(end) = xml[content_start..].find(&close) {
            let section = &xml[content_start..content_start + end];
            let mut emails = Vec::new();
            let mut pos = 0;
            while let Some(s) = section[pos..].find("<t:EmailAddress>") {
                let abs = pos + s + "<t:EmailAddress>".len();
                if let Some(e) = section[abs..].find("</t:EmailAddress>") {
                    emails.push(section[abs..abs + e].to_owned());
                    pos = abs + e;
                } else {
                    break;
                }
            }
            return emails.join(", ");
        }
    }
    String::new()
}

// ─── Tests ───────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_xml_text_works() {
        let xml = r#"<t:Subject>Hello World</t:Subject>"#;
        assert_eq!(extract_xml_text(xml, "Subject"), Some("Hello World"));
    }

    #[test]
    fn extract_attr_works() {
        let xml = r#"<t:ItemId Id="abc123" ChangeKey="xyz"/>"#;
        assert_eq!(extract_attr(xml, "ItemId", "Id"), Some("abc123"));
        assert_eq!(extract_attr(xml, "ItemId", "ChangeKey"), Some("xyz"));
    }

    #[test]
    fn parse_find_items_empty() {
        let xml = r#"<soap:Envelope><soap:Body><m:FindItemResponse></m:FindItemResponse></soap:Body></soap:Envelope>"#;
        let result = parse_find_items_response(xml).unwrap();
        assert!(result.is_empty());
    }
}
