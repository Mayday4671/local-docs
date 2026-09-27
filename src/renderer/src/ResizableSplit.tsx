import { useId, useLayoutEffect, useRef, useState, type ReactNode } from 'react'

const dividerWidth = 8
const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value))

function savedRatio(preferenceKey: string, fallback: number): number {
  try {
    const value = Number(localStorage.getItem(preferenceKey))
    if (Number.isFinite(value) && value >= 0.2 && value <= 0.8) return value
  } catch {
    // A blocked preference store must not prevent editing.
  }
  return fallback
}

function limits(width: number, minLeft: number, minRight: number) {
  const available = Math.max(1, width - dividerWidth)
  const min = Math.max(0.2, minLeft / available)
  const max = Math.min(0.8, 1 - minRight / available)
  return min <= max
    ? { min, max }
    : { min: minLeft / (minLeft + minRight), max: minLeft / (minLeft + minRight) }
}

/** This preference describes this computer's layout, not document content. */
export function ResizableSplit({
  children,
  preferenceKey,
  defaultRatio = 0.5,
  minLeft = 280,
  minRight = 280,
  leftLabel,
  rightLabel,
  className = '',
  secondaryVisible = true,
  resetLabel = '默认比例',
}: {
  children: [ReactNode, ReactNode]
  preferenceKey: string
  defaultRatio?: number
  minLeft?: number
  minRight?: number
  leftLabel: string
  rightLabel: string
  className?: string
  secondaryVisible?: boolean
  resetLabel?: string
}) {
  const root = useRef<HTMLDivElement>(null)
  const paneId = useId()
  const [ratio, setRatio] = useState(() => savedRatio(preferenceKey, defaultRatio))
  const latestRatio = useRef(ratio)
  const [width, setWidth] = useState(0)
  const [dragging, setDragging] = useState(false)
  const drag = useRef<{ pointer: number; offset: number; start: number } | null>(null)
  const { min, max } = limits(width, minLeft, minRight)
  const effective = clamp(ratio, min, max)

  useLayoutEffect(() => {
    const element = root.current!
    const measure = () => setWidth(element.getBoundingClientRect().width)
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(element)
    return () => observer.disconnect()
  }, [])

  const persist = (value: number) => {
    try {
      localStorage.setItem(preferenceKey, String(value))
    } catch {
      // Resizing remains available when layout preferences cannot be written.
    }
  }
  const update = (value: number, save = false) => {
    latestRatio.current = value
    setRatio(value)
    if (save) persist(value)
  }
  const finish = (save: boolean) => {
    if (!drag.current) return
    const start = drag.current.start
    drag.current = null
    setDragging(false)
    if (save) persist(latestRatio.current)
    else update(start)
  }

  return (
    <div
      ref={root}
      className={`resizable-split ${className}${dragging ? ' is-resizing' : ''}`}
      style={{
        gridTemplateColumns: secondaryVisible
          ? `minmax(0, ${effective}fr) ${dividerWidth}px minmax(0, ${1 - effective}fr)`
          : 'minmax(0, 1fr)',
      }}
    >
      <div id={paneId} className="split-pane">
        {children[0]}
      </div>
      {secondaryVisible && (
        <div
          className="split-divider"
          role="separator"
          tabIndex={0}
          aria-label={`调整${leftLabel}与${rightLabel}宽度`}
          aria-orientation="vertical"
          aria-controls={paneId}
          aria-valuemin={Math.round(min * 100)}
          aria-valuemax={Math.round(max * 100)}
          aria-valuenow={Math.round(effective * 100)}
          aria-valuetext={`${leftLabel} ${Math.round(effective * 100)}%，${rightLabel} ${Math.round((1 - effective) * 100)}%`}
          title={`拖动调整宽度；双击或 Enter 恢复${resetLabel}；左右方向键微调`}
          onPointerDown={(event) => {
            if (event.button !== 0 || !event.isPrimary) return
            event.preventDefault()
            event.currentTarget.focus()
            const rect = event.currentTarget.getBoundingClientRect()
            drag.current = {
              pointer: event.pointerId,
              offset: event.clientX - (rect.left + rect.width / 2),
              start: ratio,
            }
            latestRatio.current = ratio
            event.currentTarget.setPointerCapture(event.pointerId)
            setDragging(true)
          }}
          onPointerMove={(event) => {
            if (drag.current?.pointer !== event.pointerId) return
            const rect = root.current!.getBoundingClientRect()
            const bounds = limits(rect.width, minLeft, minRight)
            update(
              clamp(
                (event.clientX - drag.current.offset - rect.left - dividerWidth / 2) /
                  Math.max(1, rect.width - dividerWidth),
                bounds.min,
                bounds.max,
              ),
            )
          }}
          onPointerUp={(event) => {
            if (drag.current?.pointer !== event.pointerId) return
            finish(true)
            if (event.currentTarget.hasPointerCapture(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId)
          }}
          onPointerCancel={() => finish(false)}
          onLostPointerCapture={() => finish(true)}
          onDoubleClick={() => update(defaultRatio, true)}
          onKeyDown={(event) => {
            if (event.key === 'Escape' && drag.current) {
              event.preventDefault()
              const pointer = drag.current.pointer
              finish(false)
              if (event.currentTarget.hasPointerCapture(pointer))
                event.currentTarget.releasePointerCapture(pointer)
              return
            }
            const step = event.shiftKey ? 0.1 : 0.02
            const next =
              event.key === 'ArrowLeft'
                ? effective - step
                : event.key === 'ArrowRight'
                  ? effective + step
                  : event.key === 'Home'
                    ? min
                    : event.key === 'End'
                      ? max
                      : event.key === 'Enter'
                        ? defaultRatio
                        : null
            if (next === null) return
            event.preventDefault()
            update(clamp(next, min, max), true)
          }}
        />
      )}
      {secondaryVisible && <div className="split-pane">{children[1]}</div>}
    </div>
  )
}
