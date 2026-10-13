/** Native document commands available to external hosts. Service/second-agent
 * tools stay with the embedding application; all editing uses executeTool. */
export const DOCUMENT_COMMANDS = [
  'read_document', 'describe_tools', 'read_blocks', 'insert_content', 'replace_blocks',
  'replace_selection', 'apply_ops', 'read_revisions', 'accept_changes', 'reject_changes',
  'read_comments', 'reply_comment', 'resolve_comment', 'add_comment', 'delete_comment',
  'insert_footnote', 'insert_endnote', 'delete_note', 'edit_note', 'read_notes',
  'insert_image', 'insert_picture', 'insert_chart', 'edit_chart', 'set_header_footer',
  'set_page_setup', 'insert_section_break', 'define_style', 'list_styles', 'set_watermark',
  'insert_text_box', 'save_document', 'undo', 'redo',
] as const
export type DocumentCommand = typeof DOCUMENT_COMMANDS[number]
export const DOCUMENT_READS = new Set<string>([
  'read_document', 'describe_tools', 'read_blocks', 'read_revisions', 'read_comments',
  'read_notes', 'list_styles',
])
