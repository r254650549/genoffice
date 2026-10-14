import type { SlidesApi } from '../shared/ipc'
import type { RenderSlide } from '@genoffice/pptx-render'
export interface PresentationControls {
  flush(): Promise<void>
  save(): Promise<boolean>
  selection(): { slideIndex: number; sourceIds: string[] }
  apply(state: { slides: RenderSlide[]; dirty: boolean }): void
}
export interface OfficeHost extends SlidesApi {
  kind?: 'electron' | 'web'
  readOnly?: boolean
  bindPresentation?(controls: PresentationControls): void
  openAgent?(prompt?: string): void
  draftDirty?(): boolean
}
let host: OfficeHost
export function configureOfficeHost(value: OfficeHost) {
  host = value
}
export function getOfficeHost(): OfficeHost {
  return host
}
