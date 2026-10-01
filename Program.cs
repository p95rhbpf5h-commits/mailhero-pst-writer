/* mailhero-pst-writer — Unicode PST writer helper for MailHero (macOS).
 *
 * Copyright (C) 2026 Drake Allegrini.
 * Built on PSTFileFormat / ContinuMail, Copyright (C) 2012-2017 ROM Knowledgeware,
 * maintained by Tal Aloni.
 *
 * This program is free software: you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or (at your
 * option) any later version. See COPYING.LESSER and COPYING in this
 * repository for the full text.
 */
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using System.Net.Mail;
using System.Text.RegularExpressions;
using PSTFileFormat;

if (args.Length is < 1 or > 2)
    throw new ArgumentException("Usage: mailhero-pst-writer OUTPUT.pst [HELPER_RAM_BYTES]");

var outputPath = args[0];
// OST export is deferred past v1. This helper always writes a "SM" (PST)
// header; an OST is an account-synchronized Outlook cache and is not
// produced by renaming a PST. Refuse the extension so the app cannot
// accidentally publish a mislabeled store.
if (string.Equals(Path.GetExtension(outputPath), ".ost", StringComparison.OrdinalIgnoreCase))
    throw new ArgumentException("OST output is not supported; choose a .pst destination.");
Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
PSTFile.CreateEmptyStore(outputPath);

var file = new PSTFile(outputPath, FileAccess.ReadWrite, WriterCompatibilityMode.Outlook2007RTM);
try
{
    file.BeginSavingChanges();
    var folders = new Dictionary<string, PSTFolder>(StringComparer.OrdinalIgnoreCase);

    string? line;
    while ((line = Console.In.ReadLine()) is not null)
    {
        if (string.IsNullOrWhiteSpace(line))
            continue;

        var item = JsonSerializer.Deserialize<InputMessage>(line, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("Invalid PST writer input.");
        var folder = GetFolder(file, folders, item.FolderPath);
        var note = Note.CreateNewNote(file, folder.NodeID);
        note.Subject = item.Subject ?? string.Empty;
        note.Body = item.Body ?? string.Empty;
        if (!string.IsNullOrEmpty(item.HtmlBody))
        {
            note.PC.SetBytesProperty(PropertyID.PidTagHtml, Encoding.UTF8.GetBytes(item.HtmlBody));
            note.PC.SetInt32Property(PropertyID.PidTagNativeBody, 3);
            note.InternetCodepage = 65001;
        }
        else if (!string.IsNullOrWhiteSpace(item.BodyRtfBase64))
        {
            try
            {
                note.PC.SetBytesProperty(PropertyID.PidTagRtfCompressed,
                    Convert.FromBase64String(item.BodyRtfBase64));
                note.PC.SetInt32Property(PropertyID.PidTagNativeBody, 2);
            }
            catch (FormatException ex)
            {
                throw new InvalidDataException("Invalid Base64 RTF body.", ex);
            }
        }
        SetSender(note, item);
        SetRecipients(note, item);
        if (!string.IsNullOrWhiteSpace(item.TransportHeaders))
            note.PC.SetStringProperty((PropertyID)0x007D,
                NormalizeTransportHeaders(item.TransportHeaders, item.Subject ?? string.Empty));
        foreach (var attachment in item.Attachments ?? Array.Empty<InputAttachment>())
            WriteAttachment(file, note, attachment);
        if (!string.IsNullOrEmpty(item.MessageId))
            note.PC.SetStringProperty(PropertyID.PidTagInternetMessageId, item.MessageId);
        if (item.DeliveryTime is { } delivery)
        {
            var date = DateTime.FromFileTimeUtc(delivery);
            note.ClientSubmitTime = date;
            note.MessageDeliveryTime = date;
        }
        note.SaveChanges();
        folder.AddMessage(note);
    }

    foreach (var folder in folders.Values)
        folder.SaveChanges();
    file.EndSavingChanges();
}
finally
{
    file.CloseFile();
}

static PSTFolder GetFolder(PSTFile file, Dictionary<string, PSTFolder> folders, string[]? path)
{
    var parts = path is { Length: > 0 } ? path.Where(p => !string.IsNullOrWhiteSpace(p)).ToArray() : new[] { "Inbox" };
    var key = string.Join("/", parts);
    if (folders.TryGetValue(key, out var existing))
        return existing;

    var current = file.TopOfPersonalFolders;
    foreach (var part in parts)
    {
        var child = current.FindChildFolder(part);
        current = child ?? current.CreateChildFolder(part, FolderItemTypeName.Note);
    }
    folders[key] = current;
    return current;
}

static void SetSender(Note note, InputMessage item)
{
    if (string.IsNullOrWhiteSpace(item.SenderEmail) && string.IsNullOrWhiteSpace(item.SenderName))
        return;
    var senderText = item.SenderName ?? item.SenderEmail ?? string.Empty;
    var parsed = ParseAddress(senderText);
    var name = parsed.DisplayName;
    var email = parsed.Email.Contains('@', StringComparison.Ordinal) ? parsed.Email : (item.SenderEmail ?? parsed.Email);
    if (string.IsNullOrWhiteSpace(name) || string.Equals(name, email, StringComparison.OrdinalIgnoreCase))
        name = email;
    note.SenderName = name;
    note.SentRepresentingName = name;
    if (!string.IsNullOrWhiteSpace(email))
    {
        note.SenderAddressType = "SMTP";
        note.SenderEmailAddress = email;
        note.SentRepresentingAddressType = "SMTP";
        note.SentRepresentingEmailAddress = email;
    }
}

static void SetRecipients(Note note, InputMessage item)
{
    var recipients = new List<MessageRecipient>();
    Add(item.To, RecipientType.To, recipients);
    Add(item.Cc, RecipientType.Cc, recipients);
    Add(item.Bcc, RecipientType.Bcc, recipients);
    if (recipients.Count > 0)
        note.AddRecipients(recipients);

    static void Add(string[]? values, RecipientType type, List<MessageRecipient> result)
    {
        if (values is null) return;
        foreach (var value in values.Where(v => !string.IsNullOrWhiteSpace(v)))
        {
            var address = ParseAddress(value);
            result.Add(new MessageRecipient(address.DisplayName, address.Email, false, type));
        }
    }
}

static string NormalizeTransportHeaders(string headers, string subject)
{
    // PidTagTransportMessageHeaders and PidTagSubject must agree. MailHero's
    // PST reader intentionally prefers the transport Subject header when it
    // is present, so preserve every header while making that one canonical.
    const string pattern = @"(?im)^Subject:[^\r\n]*(?:\r?\n[ \t][^\r\n]*)*";
    return Regex.Replace(headers, pattern, "Subject: " + subject);
}

static (string DisplayName, string Email) ParseAddress(string value)
{
    value = value.Trim();
    if (value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
        value = value["mailto:".Length..].Trim();

    // MailAddress handles quoted display names, quoted local parts, comments,
    // plus-addressing, and Unicode display names. Keep the original text only
    // as a fallback because malformed but recoverable source headers are common.
    try
    {
        var parsed = new MailAddress(value);
        var displayName = parsed.DisplayName.Trim();
        return (string.IsNullOrEmpty(displayName) ? parsed.Address : displayName, parsed.Address);
    }
    catch (FormatException)
    {
        // Fall through to the deliberately conservative angle-bracket parser.
    }

    var open = value.LastIndexOf('<');
    var close = value.LastIndexOf('>');
    if (open >= 0 && close > open)
    {
        var name = value[..open].Trim().Trim('"');
        var email = value[(open + 1)..close].Trim();
        return (string.IsNullOrEmpty(name) ? email : name, email);
    }
    return (value.Trim(), value.Trim());
}

static void WriteAttachment(PSTFile file, Note note, InputAttachment input)
{
    if (string.IsNullOrWhiteSpace(input.Filename))
        throw new InvalidDataException("PST attachment filename is required.");

    byte[] payload;
    try
    {
        payload = Convert.FromBase64String(input.DataBase64 ?? string.Empty);
    }
    catch (FormatException ex)
    {
        throw new InvalidDataException($"Invalid Base64 attachment: {input.Filename}", ex);
    }

    note.CreateSubnodeBTreeIfNotExist();
    var attachment = AttachmentObject.CreateNewAttachmentObject(file, note.SubnodeBTree);
    var mimeType = string.IsNullOrWhiteSpace(input.MimeType) ? "application/octet-stream" : input.MimeType;
    attachment.PC.SetStringProperty(PropertyID.PidTagAttachLongFilename, input.Filename);
    attachment.PC.SetStringProperty(PropertyID.PidTagAttachFilename, input.Filename);
    attachment.PC.SetStringProperty(PropertyID.PidTagDisplayName, input.Filename);
    attachment.PC.SetStringProperty(PropertyID.PidTagAttachExtension, Path.GetExtension(input.Filename));
    attachment.PC.SetStringProperty(PropertyID.PidTagAttachMimeTag, mimeType);
    attachment.PC.SetInt32Property(PropertyID.PidTagAttachMethod, (int)AttachMethod.ByValue);
    attachment.PC.SetInt32Property(PropertyID.PidTagRenderingPosition, -1);
    attachment.PC.SetInt32Property(PropertyID.PidTagAttachSize, 0);
    if (!string.IsNullOrWhiteSpace(input.ContentId))
        attachment.PC.SetStringProperty(PropertyID.PidTagAttachContentId, input.ContentId);
    if (!string.IsNullOrWhiteSpace(input.ContentLocation))
        attachment.PC.SetStringProperty(PropertyID.PidTagAttachContentLocation, input.ContentLocation);
    if (input.IsInline)
    {
        attachment.PC.SetInt32Property(PropertyID.PidTagAttachFlags, 4);
        attachment.PC.SetBooleanProperty(PropertyID.PidTagAttachmentHidden, true);
    }
    attachment.PC.SetBytesProperty(PropertyID.PidTagAttachData, payload);
    var objectSize = attachment.PC.GetTotalLengthOfAllProperties() + payload.Length;
    if (objectSize > int.MaxValue)
        throw new InvalidDataException($"Attachment is too large: {input.Filename}");
    attachment.PC.SetInt32Property(PropertyID.PidTagAttachSize, objectSize);
    attachment.SaveChanges(note.SubnodeBTree);
    note.AddAttachment(attachment);
}

sealed class InputMessage
{
    [JsonPropertyName("subject")] public string? Subject { get; set; }
    [JsonPropertyName("sender_name")] public string? SenderName { get; set; }
    [JsonPropertyName("sender_email")] public string? SenderEmail { get; set; }
    [JsonPropertyName("to")] public string[]? To { get; set; }
    [JsonPropertyName("cc")] public string[]? Cc { get; set; }
    [JsonPropertyName("bcc")] public string[]? Bcc { get; set; }
    [JsonPropertyName("body")] public string? Body { get; set; }
    [JsonPropertyName("html_body")] public string? HtmlBody { get; set; }
    [JsonPropertyName("body_rtf_base64")] public string? BodyRtfBase64 { get; set; }
    [JsonPropertyName("message_id")] public string? MessageId { get; set; }
    [JsonPropertyName("delivery_time")] public long? DeliveryTime { get; set; }
    [JsonPropertyName("folder_path")] public string[]? FolderPath { get; set; }
    [JsonPropertyName("transport_headers")] public string? TransportHeaders { get; set; }
    [JsonPropertyName("attachments")] public InputAttachment[]? Attachments { get; set; }
}

sealed class InputAttachment
{
    [JsonPropertyName("filename")] public string? Filename { get; set; }
    [JsonPropertyName("mime_type")] public string? MimeType { get; set; }
    [JsonPropertyName("content_id")] public string? ContentId { get; set; }
    [JsonPropertyName("content_location")] public string? ContentLocation { get; set; }
    [JsonPropertyName("data_base64")] public string? DataBase64 { get; set; }
    [JsonPropertyName("is_inline")] public bool IsInline { get; set; }
}
