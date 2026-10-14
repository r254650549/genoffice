/** Explicit Node ports for the original native handlers. Uploaded bytes never become host paths. */
import { imageSizeFromBytes } from '@genoffice/pptx-render'
import type { SlidesDocumentPorts } from './document-handlers'

export interface PresentationAsset {
  name: string
  base64: string
}
export function createNodeDocumentPorts() {
  let file: { path: string; bytes: Buffer } | undefined
  let text = '',
    html = ''
  let image: Buffer = Buffer.alloc(0)
  const buffers = new Map<string, Buffer>()
  const bytesFor = (path: string) => {
    if (!file || path !== file.path) throw new Error('PPTX_ASSET_REQUIRED')
    return Buffer.from(file.bytes)
  }
  const decoded = (bytes: Buffer) => ({
    getSize: () => imageSizeFromBytes(bytes) ?? { width: 0, height: 0 },
    isEmpty: () => !imageSizeFromBytes(bytes),
    toPNG: () => {
      if (!bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])))
        throw new Error('PPTX_PNG_REQUIRED')
      return Buffer.from(bytes)
    },
  })
  const ports = {
    dialog: {
      showMessageBox: async () => {
        throw new Error('PPTX_OPERATION_UNSUPPORTED')
      },
    },
    dialogParent: () => undefined,
    showOpenDialogWithMemory: async () =>
      file ? { canceled: false, filePaths: [file.path] } : { canceled: true, filePaths: [] },
    readFile: async (path: string) => bytesFor(path),
    nativeImage: {
      createFromBuffer: decoded,
      createFromPath: (path: string) => decoded(bytesFor(path)),
      createThumbnailFromPath: async () => {
        throw new Error('PPTX_THUMBNAIL_UNAVAILABLE')
      },
    },
    clipboard: {
      readBuffer: (format: string) => buffers.get(format) ?? Buffer.alloc(0),
      writeBuffer: (format: string, value: Buffer) => {
        buffers.set(format, Buffer.from(value))
      },
      readText: () => text,
      readHTML: () => html,
      writeText: (value: string) => {
        text = value
        html = ''
        image = Buffer.alloc(0)
        buffers.clear()
      },
      readImage: () => decoded(image),
      write: (value: { html?: string; image?: { toPNG(): Buffer } }) => {
        html = value.html ?? ''
        image = value.image?.toPNG() ?? Buffer.alloc(0)
      },
      availableFormats: () => [...buffers.keys(), ...(image.length ? ['image/png'] : [])],
    },
    commentAuthorName: () => 'kompagent',
  } as unknown as SlidesDocumentPorts
  return {
    ports,
    prepare(asset?: PresentationAsset) {
      file = undefined
      if (!asset) return
      if (!/^[^\\/\x00-\x1f]{1,255}$/.test(asset.name) || asset.base64.length > 90 * 1024 * 1024)
        throw new Error('PPTX_ASSET_INVALID')
      const bytes = Buffer.from(asset.base64, 'base64')
      if (!bytes.length || bytes.length > 64 * 1024 * 1024) throw new Error('PPTX_ASSET_INVALID')
      file = { path: '/office-upload/' + asset.name, bytes }
    },
  }
}
