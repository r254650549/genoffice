import { IRenderManagerService } from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

/// Grows a worksheet's grid as the viewport approaches its edge.
///
/// The grid a sheet is created with is sized to its data (see
/// MINIMUM_SHEET_COLUMN_COUNT / MINIMUM_SHEET_ROW_COUNT in univer-sync), so a
/// blank workbook stops being scrollable a few columns past the last cell. This
/// grows it while the user works instead, the way Excel does: scrolling right
/// keeps revealing columns, because reaching the edge extends the grid before
/// the next scroll would have stopped at it.
///
/// The growth is a view concern only, so it must stay out of both histories:
/// it writes no journal entry and never reaches the snapshot, so the grid a
/// file declares on save is still the used range — the same split the engine
/// already makes for a <dimension> that runs to the last column (workbook.rs
/// declared_extent) — and it is dispatched as a mutation, so it pushes no undo
/// entry either. Scrolling is a view action: a pristine blank sheet that grew
/// past column F on a scroll must leave the Undo button disabled, or the first
/// Ctrl+Z after typing at AA1 shrinks the grid back instead of undoing the
/// edit.

/// Univer's last valid column and row index (0-based). The .xlsx grid caps at
/// 16384 x 1048576, so growth stops here rather than past it.
const LAST_COLUMN_INDEX = 16_383
const LAST_ROW_INDEX = 1_048_575

/// How close the viewport's first column/row gets to the edge before the grid
/// extends. The extension is only visible once the user reaches it, so a
/// viewport parked at the edge would otherwise sit against a dead border.
const COLUMN_LOOKAHEAD = 20
const ROW_LOOKAHEAD = 200

/// How much one extension adds. A column is a fraction of a screen wide and a
/// row is a couple of lines tall, so the axes need different steps to cover the
/// same distance.
const COLUMN_STEP = 26
const ROW_STEP = 200

export interface GridEdge {
  /// first visible column and row, 0-based (Univer's own indices)
  readonly startColumn: number
  readonly startRow: number
  /// Last visible indices account for wide windows, zoom and narrow columns.
  readonly endColumn?: number
  readonly endRow?: number
  /// the grid's present size, in columns and rows
  readonly columnCount: number
  readonly rowCount: number
}

export interface GridGrowth {
  readonly columnCount: number
  readonly rowCount: number
}

/**
 * The grid size the viewport at `edge` needs, or null when it already has one.
 *
 * Growth is self-limiting: the caller's new size puts the viewport back outside
 * the lookahead, so a burst of scroll events extends the grid once rather than
 * once per event.
 */
export function nextGridGrowth(edge: GridEdge): GridGrowth | null {
  let columnCount = edge.columnCount
  if (
    columnCount < LAST_COLUMN_INDEX + 1 &&
    (edge.endColumn ?? edge.startColumn) + COLUMN_LOOKAHEAD >= columnCount
  ) {
    columnCount = Math.min(columnCount + COLUMN_STEP, LAST_COLUMN_INDEX + 1)
  }
  let rowCount = edge.rowCount
  if (
    rowCount < LAST_ROW_INDEX + 1 &&
    (edge.endRow ?? edge.startRow) + ROW_LOOKAHEAD >= rowCount
  ) {
    rowCount = Math.min(rowCount + ROW_STEP, LAST_ROW_INDEX + 1)
  }
  if (columnCount === edge.columnCount && rowCount === edge.rowCount) return null
  return { columnCount, rowCount }
}

interface ScrollCommandParams {
  readonly sheetViewStartColumn?: number
  readonly sheetViewStartRow?: number
}

interface SkeletonManagerLike {
  reCalculate(): void
}

interface RenderUnitLike {
  with<T>(token: unknown): T
}

interface InjectorLike {
  get<T>(token: unknown): T
}

/** Univer's scroll command: the one carrying the viewport's first row/column. */
const SET_SCROLL_COMMAND = 'sheet.operation.set-scroll'
const SET_ACTIVE_SHEET_OPERATION = 'sheet.operation.set-worksheet-active'

/**
 * Grows the active sheet's grid as its viewport nears the edge.
 *
 * The viewport arrives through Univer's scroll command rather than a scroll
 * service: SheetScrollManagerService is declared in @univerjs/sheets-ui but is
 * not registered in this build, so asking the injector for it throws. The
 * command is dispatched on every scroll — wheel, scrollbar, keyboard — and
 * carries sheetViewStartColumn/Row directly, so no pixel-to-index conversion
 * is needed.
 *
 * Two triggers, because neither alone covers everything:
 *   - the command subscription, which is the real signal
 *   - a slow poll, for a scroll the command path does not cover, and because
 *     the render unit only exists a moment after install
 *
 * The poll reads the actual visible range, including its right/bottom edge,
 * so initial layout, resizing and zooming also grow the grid without a scroll.
 *
 * Returns a disposer.
 */
export function installGridGrowth(runtime: UniverRuntime): () => void {
  let disposed = false
  // The render unit appears after install, so resolve it per use and cache only
  // once it resolves: a cached miss would pin the growth off forever.
  let skeletonManager: SkeletonManagerLike | null = null
  let skeletonUnitId: string | null = null

  const currentSkeletonManager = (): SkeletonManagerLike | null => {
    try {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      if (!workbook) return null
      if (skeletonManager && skeletonUnitId === workbook.getId()) return skeletonManager
      const injector = (
        runtime.univer as unknown as { __getInjector(): InjectorLike }
      ).__getInjector()
      const render = injector
        .get<{ getRenderById(id: string): RenderUnitLike | null }>(IRenderManagerService)
        .getRenderById(workbook.getId())
      if (!render) return null
      skeletonManager = render.with<SkeletonManagerLike>(SheetSkeletonManagerService)
      skeletonUnitId = workbook.getId()
      return skeletonManager
    } catch {
      return null
    }
  }

  const growIfNearEdge = (startColumn: number, startRow: number): void => {
    if (disposed) return
    const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!worksheet) return
    // The first visible column alone misses a viewport wider than the entire
    // initial grid: it remains at A and cannot scroll far enough to grow.
    const visible = worksheet.getVisibleRange?.()
    const growth = nextGridGrowth({
      startColumn,
      startRow,
      endColumn: Math.max(startColumn, visible?.endColumn ?? startColumn),
      endRow: Math.max(startRow, visible?.endRow ?? startRow),
      columnCount: worksheet.getMaxColumns(),
      rowCount: worksheet.getMaxRows(),
    })
    if (!growth) return
    // Dispatch the mutations directly. The facade's setColumnCount/setRowCount
    // run SetWorksheetColumnCountCommand / SetWorksheetRowCountCommand, and both
    // handlers call undoRedoService.pushUndoRedo, so every lookahead step would
    // otherwise leave an undo step behind. The mutation handlers call
    // worksheet.setColumnCount / setRowCount and stop there, so the grid grows
    // with the undo stack untouched.
    const unitId = worksheet.getSheet().getUnitId()
    const subUnitId = worksheet.getSheetId()
    if (growth.columnCount !== worksheet.getMaxColumns()) {
      runtime.univerAPI.syncExecuteCommand('sheet.mutation.set-worksheet-column-count', {
        unitId,
        subUnitId,
        columnCount: growth.columnCount,
      })
    }
    if (growth.rowCount !== worksheet.getMaxRows()) {
      runtime.univerAPI.syncExecuteCommand('sheet.mutation.set-worksheet-row-count', {
        unitId,
        subUnitId,
        rowCount: growth.rowCount,
      })
    }
    // the skeleton caches the row/column sizes it was built with
    currentSkeletonManager()?.reCalculate()
  }

  let lastColumn = 0
  let lastRow = 0

  const onCommand = runtime.univerAPI.onCommandExecuted(
    (command: { id?: string; params?: unknown }) => {
      if (command?.id === SET_ACTIVE_SHEET_OPERATION) {
        // The viewport position belongs to the sheet it was scrolled on.
        lastColumn = 0
        lastRow = 0
        return
      }
      if (command?.id !== SET_SCROLL_COMMAND) return
      const params = command.params as ScrollCommandParams | undefined
      if (!params) return
      lastColumn = params.sheetViewStartColumn ?? lastColumn
      lastRow = params.sheetViewStartRow ?? lastRow
      growIfNearEdge(lastColumn, lastRow)
    },
  )

  // Backstop: a viewport move the command stream does not carry, and the
  // window between install and the render unit existing.
  const timer = window.setInterval(() => {
    if (disposed) return
    const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!worksheet) return
    const visible = worksheet.getVisibleRange?.()
    if (!visible) return
    growIfNearEdge(visible.startColumn, visible.startRow)
  }, 250)

  return () => {
    disposed = true
    window.clearInterval(timer)
    onCommand?.dispose?.()
  }
}
