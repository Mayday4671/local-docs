import { Select } from './Select'

export function ZoomSelect({
  label,
  value,
  onChange,
  disabled = false,
}: {
  label: string
  value: number
  onChange: (value: number) => void
  disabled?: boolean
}) {
  const choices = [...new Set([50, 75, 90, 100, 125, 150, 175, 200, 250, 300, value])].sort(
    (a, b) => a - b,
  )
  return (
    <Select
      className="zoom-select"
      aria-label={label}
      title="Ctrl + 滚轮：放大 / 缩小"
      value={value}
      disabled={disabled}
      onChange={(event) => onChange(Number(event.target.value))}
    >
      {choices.map((zoom) => (
        <option key={zoom} value={zoom}>
          {zoom}%
        </option>
      ))}
    </Select>
  )
}
