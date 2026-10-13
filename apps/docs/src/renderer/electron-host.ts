import { configureOfficeHost } from './office-host'
configureOfficeHost({ ...window.desktop, kind: 'electron' })
