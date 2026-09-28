import { useEffect, useState } from 'react'
import { Download } from 'lucide-react'

export function FileDropZone({
  disabled,
  destination,
  onFiles,
  onBlocked,
}: {
  disabled: boolean
  destination: string
  onFiles: (files: File[]) => void
  onBlocked: () => void
}) {
  const [over, setOver] = useState(false)
  useEffect(() => {
    let depth = 0
    const hasFiles = (e: DragEvent) => [...(e.dataTransfer?.types ?? [])].includes('Files')
    const enter = (e: DragEvent) => {
      if (hasFiles(e)) {
        e.preventDefault()
        depth++
        setOver(true)
      }
    }
    const move = (e: DragEvent) => {
      if (!hasFiles(e)) return
      e.preventDefault()
      if (e.dataTransfer) e.dataTransfer.dropEffect = disabled ? 'none' : 'copy'
    }
    const leave = (e: DragEvent) => {
      if (hasFiles(e) && --depth <= 0) {
        depth = 0
        setOver(false)
      }
    }
    const reset = () => {
      depth = 0
      setOver(false)
    }
    const drop = (e: DragEvent) => {
      if (!hasFiles(e)) return
      // Prevent Chromium navigating to dropped documents, including in the editor.
      e.preventDefault()
      e.stopPropagation()
      reset()
      if (disabled) {
        onBlocked()
        return
      }
      const files = [...(e.dataTransfer?.files ?? [])]
      if (files.length) onFiles(files)
    }
    window.addEventListener('dragenter', enter)
    window.addEventListener('dragover', move)
    window.addEventListener('dragleave', leave)
    window.addEventListener('drop', drop, true)
    window.addEventListener('dragend', reset)
    window.addEventListener('blur', reset)
    return () => {
      window.removeEventListener('dragenter', enter)
      window.removeEventListener('dragover', move)
      window.removeEventListener('dragleave', leave)
      window.removeEventListener('drop', drop, true)
      window.removeEventListener('dragend', reset)
      window.removeEventListener('blur', reset)
    }
  }, [disabled, onFiles, onBlocked])
  return over ? (
    <div className="file-drop-overlay" role="status">
      <div>
        <Download size={38} />
        <strong>
          {disabled ? '请返回文件列表并完成当前操作后再拖入' : `松开即可添加到「${destination}」`}
        </strong>
        <span>支持一次拖入多个文件 · 保留电脑中的原文件</span>
      </div>
    </div>
  ) : null
}
