import { clipboard, nativeImage } from 'electron'
import { createElementClipboard } from './element-clipboard-core'
export { ELEMENT_CLIPBOARD_FORMAT } from './element-clipboard-core'
export type { ElementClipboardIdentity } from './element-clipboard-core'
export const {
  isElementClipboardToken,
  elementClipboardMarkerMatches,
  canWriteElementClipboardImage,
  writeElementClipboardImage,
} = createElementClipboard({ clipboard, nativeImage })
