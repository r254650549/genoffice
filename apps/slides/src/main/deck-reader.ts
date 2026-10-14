/** Original GenOffice Slides MCP model readout, shared by desktop and Web. */
import { elementDurableId, slideDurableId, type SlideElement } from '@genoffice/pptx-engine'
import type { Session } from './session-document'
const EMU_PER_PX_96 = 9525
/** per-text cap in read_deck, mirroring the CLI's describeDeck clipping */
const MAX_TEXT_CHARS = 300
/** table rows echoed per element (the CLI caps its own table output the same way) */
const MAX_TABLE_ROWS = 3

function clip(text: string, max = MAX_TEXT_CHARS): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

/**
 * Compact model readout for the agent: ids, geometry (EMU) and text per element.
 *
 * Group children carry the group's child coordinate system, not document-space
 * EMU, so their box is reported as `local` (with the parent's own box alongside)
 * instead of being passed off as an absolute position — an agent placing a
 * follow-up setTransform needs to know which frame it is addressing.
 */
function elementInfo(el: SlideElement): Record<string, unknown> {
  const t = el.transform?.offset ?? { x: 0, y: 0, cx: 0, cy: 0 }
  const info: Record<string, unknown> = {
    id: elementDurableId(el) ?? el.id,
    type: el.type,
    x: t.x,
    y: t.y,
    w: t.cx,
    h: t.cy,
  }
  if (el.type === 'text' || el.type === 'shape') {
    const paragraphs = el.text?.paragraphs ?? []
    if (paragraphs.length) {
      const full = paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
      info.text = clip(full)
      if (full.length > MAX_TEXT_CHARS) info.textTruncated = true
    }
  } else if (el.type === 'table') {
    const rows = el.rows
    const totalRows = rows.length
    const body = rows
      .slice(0, MAX_TABLE_ROWS)
      .map((row) =>
        row
          .map((cell) =>
            (cell.text?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join('')).join(''),
          )
          .join(' | '),
      )
      .join('\n')
    info.text = clip(body)
    if (totalRows > MAX_TABLE_ROWS) {
      info.textTruncated = true
      info.rows = totalRows
    }
  } else if (el.type === 'group') {
    // children are in the group's local frame: say so rather than implying
    // document-space coordinates
    info.children = el.children.map((child) => ({
      ...elementInfo(child),
      coordinates: 'local-to-group',
    }))
  }
  return info
}

export function readDeckModel(session: Session): Record<string, unknown> {
  const deck = session.opened.deck
  return {
    emuPerPx: EMU_PER_PX_96,
    note: 'x/y/w/h are EMU in slide space; elements inside a group use the group-local frame (marked coordinates:"local-to-group").',
    slideSize: {
      widthPx: Math.round(deck.size.cx / EMU_PER_PX_96),
      heightPx: Math.round(deck.size.cy / EMU_PER_PX_96),
    },
    slides: deck.slides.map((slide, index) => ({
      index,
      id: slideDurableId(slide),
      elements: slide.elements.map(elementInfo),
    })),
  }
}
