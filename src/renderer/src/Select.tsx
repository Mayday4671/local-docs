import { createElement, type ComponentProps } from 'react'

/** Native keyboard and form behavior, with a truncatable selected label. */
export function Select({ children, ...props }: ComponentProps<'select'>) {
  return (
    <select {...props}>
      <button type="button">{createElement('selectedcontent')}</button>
      {children}
    </select>
  )
}
