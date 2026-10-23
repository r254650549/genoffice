import { beforeEach, describe, expect, it, vi } from 'vitest'

const electron = vi.hoisted(() => ({
  clipboard: {
    readBuffer: vi.fn(),
    readHTML: vi.fn(),
    write: vi.fn(),
  },
  nativeImage: {
    createFromBuffer: vi.fn(),
  },
}))

vi.mock('electron', () => electron)

import {
  canWriteElementClipboardImage,
  elementClipboardMarkerMatches,
  writeElementClipboardImage,
} from '../src/main/element-clipboard'

const token = 'copy-token-1'

beforeEach(() => {
  electron.clipboard.readBuffer.mockReset().mockReturnValue(Buffer.alloc(0))
  electron.clipboard.readHTML.mockReset().mockReturnValue('')
  electron.clipboard.write.mockReset()
  electron.nativeImage.createFromBuffer.mockReset().mockImplementation((bytes: Buffer) => ({
    isEmpty: () => bytes.length === 0 || bytes[0] === 0,
    toPNG: () => Buffer.from('normalized-png'),
  }))
})

describe('element clipboard OS marker', async () => {
  it('only recognizes the exact custom-format token as internal', async () => {
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from(token))
    expect(await elementClipboardMarkerMatches(token)).toBe(true)

    electron.clipboard.readBuffer.mockReturnValue(Buffer.from(`${token}-other`))
    expect(await elementClipboardMarkerMatches(token)).toBe(false)
  })

  it('recognizes the token in the HTML image marker when custom formats are absent', async () => {
    electron.clipboard.readHTML.mockReturnValue(
      '<img src="data:image/png;base64,cG5n" data-genoffice-slides-elements="copy-token-1">',
    )

    expect(await elementClipboardMarkerMatches(token)).toBe(true)
  })

  it('rejects a stale token and an externally replaced clipboard before image write', async () => {
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from('new-token'))
    const current = { token: 'new-token', senderId: 42 }

    expect(await canWriteElementClipboardImage(current, 42, token)).toBe(false)
    expect(await canWriteElementClipboardImage(current, 7, 'new-token')).toBe(false)
    expect(await canWriteElementClipboardImage(current, 42, 'new-token')).toBe(true)

    electron.clipboard.readBuffer.mockReturnValue(Buffer.from('external-copy'))
    expect(await canWriteElementClipboardImage(current, 42, 'new-token')).toBe(false)
  })
})

describe('element clipboard PNG writer', async () => {
  it('writes normalized PNG and HTML marker in one clipboard operation', async () => {
    const decoded = { isEmpty: () => false, toPNG: () => Buffer.from('normalized-png') }
    electron.nativeImage.createFromBuffer.mockReturnValue(decoded)
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from('copy token'))

    expect(await writeElementClipboardImage('copy token', 'cG5n')).toBe(true)

    expect(electron.clipboard.write).toHaveBeenCalledOnce()
    expect(electron.clipboard.write).toHaveBeenCalledWith({
      image: decoded,
      html: '<img src="data:image/png;base64,bm9ybWFsaXplZC1wbmc=" data-genoffice-slides-elements="copy%20token">',
    })
    expect(electron.nativeImage.createFromBuffer).toHaveBeenCalledOnce()
  })

  it('does not overwrite an external copy made while PNG normalization runs', async () => {
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from(token))
    electron.nativeImage.createFromBuffer.mockReturnValue({
      isEmpty: () => false,
      toPNG: () => {
        electron.clipboard.readBuffer.mockReturnValue(Buffer.from('external-copy'))
        return Buffer.from('normalized-png')
      },
    })

    expect(await writeElementClipboardImage(token, 'cG5n')).toBe(false)
    expect(electron.clipboard.write).not.toHaveBeenCalled()
  })

  it('leaves the clipboard untouched for an empty or undecodable PNG', async () => {
    expect(await writeElementClipboardImage(token, '')).toBe(false)
    expect(await writeElementClipboardImage(token, 'AA==')).toBe(false)

    expect(electron.clipboard.write).not.toHaveBeenCalled()
  })

  it('leaves the clipboard untouched when PNG decoding throws', async () => {
    electron.nativeImage.createFromBuffer.mockImplementationOnce(() => {
      throw new Error('invalid PNG')
    })

    expect(await writeElementClipboardImage(token, 'cG5n')).toBe(false)
    expect(electron.clipboard.write).not.toHaveBeenCalled()
  })
})
