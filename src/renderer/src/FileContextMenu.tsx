import { useLayoutEffect, useRef, type ReactNode } from 'react'
import { createPortal } from 'react-dom'

export interface FileMenuAnchor {
  x: number
  y: number
  target: HTMLElement
  ids: string[]
}
export interface FileMenuItem {
  label: string
  icon: ReactNode
  action: () => void
  danger?: boolean
  divider?: boolean
}

export function FileContextMenu({
  anchor,
  title,
  items,
  onClose,
}: {
  anchor: FileMenuAnchor
  title: string
  items: FileMenuItem[]
  onClose: () => void
}) {
  const menu = useRef<HTMLDivElement>(null)
  const closeRef = useRef(onClose)
  closeRef.current = onClose
  useLayoutEffect(() => {
    const element = menu.current!
    const box = element.getBoundingClientRect()
    element.style.left = `${Math.max(8, Math.min(anchor.x, innerWidth - box.width - 8))}px`
    element.style.top = `${Math.max(8, Math.min(anchor.y, innerHeight - box.height - 8))}px`
    element.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true })
    const outside = (event: Event) => {
      if (event.target instanceof Node && !element.contains(event.target)) closeRef.current()
    }
    const dismiss = () => closeRef.current()
    window.addEventListener('pointerdown', outside)
    window.addEventListener('scroll', outside, true)
    window.addEventListener('resize', dismiss)
    window.addEventListener('blur', dismiss)
    return () => {
      window.removeEventListener('pointerdown', outside)
      window.removeEventListener('scroll', outside, true)
      window.removeEventListener('resize', dismiss)
      window.removeEventListener('blur', dismiss)
    }
  }, [anchor])
  const closeAndFocus = () => {
    onClose()
    if (anchor.target.isConnected) anchor.target.focus({ preventScroll: true })
  }
  return createPortal(
    <div
      ref={menu}
      className="popover file-context-menu"
      role="menu"
      aria-label="文件操作"
      style={{ left: anchor.x, top: anchor.y }}
      onContextMenu={(event) => event.preventDefault()}
      onKeyDown={(event) => {
        const buttons = [...menu.current!.querySelectorAll<HTMLButtonElement>('button')]
        const current = buttons.indexOf(document.activeElement as HTMLButtonElement)
        if (event.key === 'Escape' || event.key === 'Tab') {
          event.preventDefault()
          event.stopPropagation()
          closeAndFocus()
        } else if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
          event.preventDefault()
          const next =
            event.key === 'Home'
              ? 0
              : event.key === 'End'
                ? buttons.length - 1
                : (current + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length
          buttons[next]?.focus()
        }
      }}
    >
      <div className="file-menu-title" title={title}>
        {title}
      </div>
      {items.map((item) => (
        <button
          key={item.label}
          type="button"
          role="menuitem"
          tabIndex={-1}
          className={`${item.danger ? 'destructive' : ''} ${item.divider ? 'menu-divider' : ''}`}
          onClick={() => {
            closeAndFocus()
            item.action()
          }}
        >
          {item.icon}
          {item.label}
        </button>
      ))}
    </div>,
    document.body,
  )
}
