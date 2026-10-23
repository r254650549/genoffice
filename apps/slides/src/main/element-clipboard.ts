import { nativeImage } from 'electron'
import {nativeClipboardPort} from '../../../../packages/electron-utils/src/clipboard-port'
const clipboard=nativeClipboardPort()
import { createElementClipboard } from './element-clipboard-core'
export { ELEMENT_CLIPBOARD_FORMAT } from './element-clipboard-core'
export type { ElementClipboardIdentity } from './element-clipboard-core'
export const {
  isElementClipboardToken,
  elementClipboardMarkerMatches,
  canWriteElementClipboardImage,
  writeElementClipboardImage,
} = createElementClipboard({ clipboard, nativeImage })
