import { describe, expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import {
  isKompAgentOfficeHost,
  officeAgentName,
  OfficeAgentMark,
  kompagentLogoUrl,
} from '../src/office-agent-brand'

describe('Office host branding', () => {
  it('uses KompAgent only for its web adapter and managed native bridge', () => {
    for (const host of [
      { kind: 'web' as const },
      { kind: 'electron' as const, hostManagedAgent: true },
    ]) {
      expect(isKompAgentOfficeHost(host)).toBe(true)
      expect(officeAgentName(host)).toBe('KompAgent')
      const mark = OfficeAgentMark({ host, size: 26, fallback: 'native-mark' })
      expect(mark.type).toBe('img')
      expect(mark.props).toMatchObject({ src: kompagentLogoUrl, width: 26, height: 26, alt: '' })
    }
  })

  it('preserves standalone native names and marks', () => {
    for (const host of [undefined, { kind: 'electron' as const }, { hostManagedAgent: false }]) {
      expect(isKompAgentOfficeHost(host)).toBe(false)
      expect(officeAgentName(host)).toBe('Genspark AI')
      expect(officeAgentName(host, 'Genspark')).toBe('Genspark')
      expect(OfficeAgentMark({ host, size: 22, fallback: 'native-mark' }).props.children).toBe(
        'native-mark',
      )
    }
  })

  it('keeps the supplied brand asset bytes pinned to their recorded source', () => {
    const asset = readFileSync(new URL('../src/assets/kompagent-logo.png', import.meta.url))
    const provenance = JSON.parse(
      readFileSync(
        new URL('../src/assets/kompagent-logo.provenance.json', import.meta.url),
        'utf8',
      ),
    )
    expect(createHash('sha256').update(asset).digest('hex')).toBe(provenance.sha256)
    expect(provenance.source).toBe('KompAgent/assets/alpha-logo.png')
  })
})
