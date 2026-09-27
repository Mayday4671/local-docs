import { useEffect, useId, useRef, useState } from 'react'
import { Check, Monitor, Moon, Sun } from 'lucide-react'
import type { ThemePreference } from '../../shared/types'

const choices = [
  { value: 'light', label: '浅色', Icon: Sun },
  { value: 'dark', label: '深色', Icon: Moon },
  { value: 'system', label: '跟随系统', Icon: Monitor },
] as const

export function ThemeSwitcher({
  theme,
  disabled,
  onChange,
}: {
  theme: ThemePreference
  disabled: boolean
  onChange: (theme: ThemePreference) => void
}) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const id = useId()
  const selected = choices.find((choice) => choice.value === theme)!
  useEffect(() => {
    if (!open) return
    root.current?.querySelector<HTMLButtonElement>('[aria-checked="true"]')?.focus()
    const outside = (event: PointerEvent) => {
      if (event.target instanceof Node && !root.current?.contains(event.target)) setOpen(false)
    }
    document.addEventListener('pointerdown', outside)
    return () => document.removeEventListener('pointerdown', outside)
  }, [open])
  return (
    <div
      className="theme-switcher"
      ref={root}
      onKeyDown={(event) => {
        if (!open) return
        if (event.key === 'Escape') {
          event.preventDefault()
          setOpen(false)
          trigger.current?.focus()
        } else if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
          event.preventDefault()
          const buttons = [
            ...root.current!.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]'),
          ]
          const current = buttons.indexOf(document.activeElement as HTMLButtonElement)
          const next =
            event.key === 'Home'
              ? 0
              : event.key === 'End'
                ? buttons.length - 1
                : (current + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length
          buttons[next]?.focus()
        } else if (event.key === 'Tab') setOpen(false)
      }}
    >
      <button
        ref={trigger}
        className="theme-trigger"
        aria-label="切换主题"
        title={`主题：${selected.label}`}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        disabled={disabled}
        onClick={() => setOpen(!open)}
      >
        <selected.Icon size={21} />
      </button>
      {open && (
        <div id={id} className="theme-menu" role="menu" aria-label="外观主题">
          {choices.map(({ value, label, Icon }) => (
            <button
              key={value}
              role="menuitemradio"
              aria-checked={theme === value}
              onClick={() => {
                onChange(value)
                setOpen(false)
                trigger.current?.focus()
              }}
            >
              <Icon size={17} />
              <span>{label}</span>
              {theme === value && <Check size={16} />}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
