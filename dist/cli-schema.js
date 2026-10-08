/** Per-tool output schema describes native data, including optional result branches. */
const fields = {
    list_accounts: 'defaultAccount accounts[]',
    begin_account_auth: 'accountId email authUrl', finish_account_auth: 'accountId email enabled!',
    read_emails: 'accounts[] successfulAccounts[] account query emails[] count# totalFound# nextPageToken resultSizeEstimate# pagination{} errors[]',
    search_emails: 'accounts[] successfulAccounts[] account query emails[] count# totalFound# nextPageToken resultSizeEstimate# pagination{} errors[]',
    get_email_thread: 'account threadId messages[] count#', get_labels: 'account labels[] count#',
    mark_as_read: 'account messageIds[] succeededCount#', archive_emails: 'account messageIds[] succeededCount#',
    trash_emails: 'account messageIds[] succeededCount#', add_labels: 'account messageIds[] labelIds[] succeededCount#',
    remove_labels: 'account messageIds[] labelIds[] succeededCount#', create_label: 'account label{} succeededCount#',
    delete_label: 'account labelId succeededCount#', block_sender: 'account sender action filter{} succeededCount# trashedCount# warnings[] errors[]',
    list_blocked_senders: 'account filters[] count#', unblock_sender: 'account filterIds[] succeededCount# sender',
    mute_thread: 'account threadId scope thread_archived! filter_created! filter_id subject_matched warning warnings[] succeededCount#',
    create_draft: 'account draftId threadId attachmentCount# succeededCount#', delete_drafts: 'account draftIds[] succeededCount#',
    send_draft: 'account draftId messageId threadId succeededCount#', list_drafts: 'account drafts[] count#',
    search_drafts: 'account drafts[] count# query', send_email: 'account messageId threadId attachmentCount# succeededCount#',
    get_attachment: 'account messageId attachment{} succeededCount#', get_all_attachments: 'account messageId attachments[] count# succeededCount# errors[]',
    unsubscribe_from_email: 'account messageId status succeededCount# reason method url httpStatus# to subject sentMessageId threadId errors[]',
    search_drive_files: 'accounts[] successfulAccounts[] query files[] count# totalFound# errors[]',
    list_drive_files: 'accounts[] successfulAccounts[] folderId query files[] count# totalFound# nextPageToken pagination{} errors[]',
    get_drive_file: 'account file{}', get_drive_file_content: 'account fileId id filename contentType sizeBytes# isInline! savedPath text extractionMethod textTruncated! extractionError',
    upload_drive_file: 'account fileId created! file{}', create_drive_folder: 'account folderId created! folder{}',
    update_drive_file: 'account fileId updated! file{}', trash_drive_file: 'account fileId trashed! updated!',
    share_drive_file: 'account fileId permissionId role type email updated!',
    sheets_get: 'account id title url sheets[]', sheets_read: 'account spreadsheetId range values[] count#',
    sheets_write: 'account spreadsheetId updated! updatedRange updatedRows# updatedCells#',
    sheets_append: 'account spreadsheetId updated! updatedRange updatedRows#',
    sheets_create: 'account spreadsheetId created! id url', sheets_add_tab: 'account spreadsheetId created! sheetId# title',
    sheets_rename_tab: 'account spreadsheetId previousTitle title updated!', sheets_delete_tab: 'account spreadsheetId sheetTitle deleted!',
    sheets_format: 'account spreadsheetId sheetTitle range updated!', sheets_add_chart: 'account spreadsheetId sheetTitle chartType dataRange created! chartId#',
    sheets_insert_dimension: 'account spreadsheetId sheetTitle dimension startIndex# count# updated!',
    sheets_delete_dimension: 'account spreadsheetId sheetTitle dimension startIndex# count# deleted!',
    docs_get: 'account documentId title body url', docs_create: 'account documentId created! url',
    docs_append: 'account documentId updated!', docs_replace_text: 'account documentId updated! occurrencesChanged#',
    docs_insert_table: 'account documentId rows# columns# updated!', docs_apply_style: 'account documentId startIndex# endIndex# style updated!',
    list_calendars: 'accounts[] successfulAccounts[] calendars[] count# errors[]', list_events: 'accounts[] successfulAccounts[] calendarId timeMin timeMax query events[] count# totalFound# errors[]',
    get_event: 'account event{}', create_event: 'account calendarId eventId created! event{}',
    update_event: 'account calendarId eventId updated! event{}', delete_event: 'account calendarId eventId deleted!',
};
export function outputSchemaFor(name) {
    const properties = Object.fromEntries((fields[name] ?? 'account succeededCount# warnings[] errors[]').split(' ').map((entry) => {
        const match = /^(.*?)(\[\]|\{\}|#|!)?$/.exec(entry);
        const [, key, suffix] = match;
        return [key, suffix === '[]' ? { type: 'array', items: {} } : suffix === '{}' ? { type: 'object' } : { type: suffix === '#' ? 'number' : suffix === '!' ? 'boolean' : ['defaultAccount', 'text'].includes(key) ? ['string', 'null'] : 'string' }];
    }));
    return {
        type: 'object', required: ['ok', 'meta'],
        properties: {
            ok: { type: 'boolean' },
            data: { type: 'object', properties, additionalProperties: true },
            error: { type: 'object', required: ['code', 'message', 'retryable'], properties: { code: { type: 'string' }, message: { type: 'string' }, retryable: { type: 'boolean' }, recovery: { type: 'string' } } },
            meta: { type: 'object', required: ['version'], properties: { version: { const: 1 }, truncation: { type: 'array' }, errors: { type: 'array' } }, additionalProperties: true },
        }, additionalProperties: false,
    };
}
//# sourceMappingURL=cli-schema.js.map