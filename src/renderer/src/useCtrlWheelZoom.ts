import { useEffect, type Dispatch, type RefObject, type SetStateAction } from 'react'

export const MIN_DOCUMENT_ZOOM = 50
export const MAX_DOCUMENT_ZOOM = 300

/** A local non-passive listener prevents Chromium from zooming the whole application. */
export function useCtrlWheelZoom(
  root: RefObject<HTMLElement | null>,
  setZoom: Dispatch<SetStateAction<number>>,
  enabled = true,
) {
  useEffect(() => {
    const element = root.current
    if (!element || !enabled) return
    let remainder = 0,
      previousTime = 0,
      pendingSteps = 0,
      frame = 0
    const wheel = (event: WheelEvent) => {
      if (!event.ctrlKey || !event.deltaY) {
        remainder = 0
        return
      }
      event.preventDefault()
      if (element.querySelector('select:open')) return
      // Normalize pixel, line and page wheels; accumulate small trackpad deltas.
      const delta = event.deltaY / (event.deltaMode === 1 ? 3 : event.deltaMode === 2 ? 1 : 100)
      if (event.timeStamp - previousTime > 250 || Math.sign(delta) !== Math.sign(remainder))
        remainder = 0
      previousTime = event.timeStamp
      remainder += delta
      const steps = Math.trunc(remainder)
      remainder -= steps
      pendingSteps -= steps
      if (!pendingSteps || frame) return
      frame = requestAnimationFrame(() => {
        const change = pendingSteps * 10
        pendingSteps = 0
        frame = 0
        setZoom((current) =>
          Math.min(MAX_DOCUMENT_ZOOM, Math.max(MIN_DOCUMENT_ZOOM, current + change)),
        )
      })
    }
    element.addEventListener('wheel', wheel, { passive: false })
    return () => {
      element.removeEventListener('wheel', wheel)
      cancelAnimationFrame(frame)
    }
  }, [root, setZoom, enabled])
}
