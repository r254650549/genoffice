import { afterEach, expect, it, vi } from 'vitest'
import { managedOfficeHost } from '../src/managed-renderer'
afterEach(() => vi.unstubAllGlobals())
it('keeps a host-owned saved version read only for all embedded editors', async () => {
  const event = vi.fn(),
    request = vi.fn(),
    onControl = vi.fn()
  vi.stubGlobal('window', { kompagentOffice: { readOnly: true, event, request, onControl } })
  vi.stubGlobal('document', { body: { inert: false } })
  const host = managedOfficeHost({
    reportMcpResult: vi.fn(),
    documentChanged: vi.fn(),
    beforeMcpCommand: vi.fn<(value: unknown) => Promise<void>>(),
  })
  expect(host.readOnly).toBe(true)
  host.documentChanged()
  expect(event).not.toHaveBeenCalled()
  await expect(host.beforeMcpCommand({ command: 'apply_ops' })).rejects.toThrow('OFFICE_READ_ONLY')
  await expect(host.beforeMcpCommand({ command: 'read_document' })).resolves.toBeUndefined()
  expect(request).not.toHaveBeenCalled()
})
