import {managedOfficeHost} from '../../../../packages/electron-utils/src/managed-renderer'
import { configureOfficeHost } from './office-host'
configureOfficeHost(managedOfficeHost({ ...window.desktop, kind: 'electron' as const }))
