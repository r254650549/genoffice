import type { ReactNode } from 'react'
import './office-agent-brand.css'

export interface OfficeAgentBrandHost {
  kind?: 'electron' | 'web'
  hostManagedAgent?: boolean
}

// The Web adapter speaks kompagent.genoffice.host/v1. Native adapters opt in
// only when the KompAgent bridge is present; standalone GenOffice stays native.
export function isKompAgentOfficeHost(host: OfficeAgentBrandHost | undefined): boolean {
  return host?.kind === 'web' || host?.hostManagedAgent === true
}

export function officeAgentName(
  host: OfficeAgentBrandHost | undefined,
  nativeName = 'Genspark AI',
): string {
  return isKompAgentOfficeHost(host) ? 'KompAgent' : nativeName
}

export const kompagentLogoUrl = new URL('./assets/kompagent-logo.png', import.meta.url).href

export function OfficeAgentMark({
  host,
  size,
  fallback,
}: {
  host: OfficeAgentBrandHost | undefined
  size: number
  fallback: ReactNode
}) {
  if (!isKompAgentOfficeHost(host)) return <>{fallback}</>
  return (
    <img
      src={kompagentLogoUrl}
      alt=""
      aria-hidden="true"
      className="office-host-agent-mark"
      width={size}
      height={size}
      style={{ display: 'block', objectFit: 'contain' }}
    />
  )
}
