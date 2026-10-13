/** KompAgent Web fork: explicit browser/Electron boundary; no preload globals. */
import type {Lang} from '@genoffice/i18n';
export interface OfficeHost {
  copyImageToClipboard?(dataUrl: string, metadata: string): Promise<boolean>;
  onLanguageChanged(listener: (language: Lang) => void): () => void;
}
let host: OfficeHost | undefined;
export function configureOfficeHost(value: OfficeHost) { host = value; }
export function getOfficeHost() { return host; }
